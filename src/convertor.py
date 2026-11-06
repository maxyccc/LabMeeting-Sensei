"""Restore full-resolution pictures while preserving native PDF text and artwork.
Python 3.11+. Install: python -m pip install pymupdf pillow pywin32
export(ppt_path, pdf_path, base_pdf=None)

For the directly tested path, supply the normal PowerPoint-exported PDF as
base_pdf. With no base_pdf, Windows desktop PowerPoint creates that PDF first;
that automation step requires PowerPoint and cannot run on Linux/macOS.

Supports PNG/JPEG images, nested rotated/flipped groups, fractional/negative
cropping, color-key transparency, grayscale, black-and-white thresholding, and
black-anchored duotone (color ramp calibrated from the native PDF). Never upsamples
the exported image: pixels come from the embedded PPTX source. Color effects are
applied to that source at its original dimensions.

Restoration verifies each tagged PDF Form, source image dimensions, exact RGB
and alpha hashes, and crop instructions after saving. It does not depend on
get_image_info() enumeration order. Native text/page instructions are preserved.
Unsupported effects or ambiguous geometry fail instead of silently skipping.
"""
from pathlib import Path
import argparse, hashlib, io, json, os, posixpath, tempfile, zipfile
import xml.etree.ElementTree as ET
from PIL import Image, ImageChops, ImageFilter
import pymupdf as fitz

NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}


def _rels(z,part):
    folder,name=posixpath.split(part)
    path=posixpath.join(folder,'_rels',name+'.rels')
    return {r.get('Id'):posixpath.normpath(posixpath.join(folder,r.get('Target')))
            for r in ET.fromstring(z.read(path)) if r.get('TargetMode')!='External'}


def _local_transform(t, group=False):
    o=t.find('a:off',NS);e=t.find('a:ext',NS)
    x,y=[int(o.get(k))/12700 for k in ('x','y')]
    w,h=[int(e.get(k))/12700 for k in ('cx','cy')]
    if group:
        co=t.find('a:chOff',NS);ce=t.find('a:chExt',NS)
        cw,ch=[int(ce.get(k))/12700 for k in ('cx','cy')]
        cx,cy=[int(co.get(k))/12700 for k in ('x','y')]
        sx=w/cw if cw else 1;sy=h/ch if ch else 1
        m=fitz.Matrix(sx,0,0,sy,x-sx*cx,y-sy*cy)
    else:m=fitz.Matrix(w,0,0,h,x,y)
    fx=-1 if t.get('flipH') in ('1','true') else 1
    fy=-1 if t.get('flipV') in ('1','true') else 1
    angle=int(t.get('rot',0))/60000
    center=(x+w/2,y+h/2)
    about=(fitz.Matrix(1,0,0,1,-center[0],-center[1]) *
           fitz.Matrix(fx,fy) * fitz.Matrix(angle) *
           fitz.Matrix(1,0,0,1,center[0],center[1]))
    return m*about


def _read(pptx):
    slides=[]
    with zipfile.ZipFile(pptx) as z:
        pres=ET.fromstring(z.read('ppt/presentation.xml'));rels=_rels(z,'ppt/presentation.xml')
        sz=pres.find('p:sldSz',NS);size=tuple(int(sz.get(k))/12700 for k in ('cx','cy'))
        for sid in pres.findall('p:sldIdLst/p:sldId',NS):
            part=rels[sid.get('{'+NS['r']+'}id')];sr=_rels(z,part);pictures=[]
            def walk(tree,parent):
                for node in tree:
                    tag=node.tag.split('}')[-1]
                    if tag=='grpSp':
                        if not node.findall('.//p:pic',NS):continue
                        t=node.find('p:grpSpPr/a:xfrm',NS)
                        walk(node,_local_transform(t,True)*parent)
                    elif tag=='pic':
                        name=node.find('p:nvPicPr/p:cNvPr',NS).get('name','Picture')
                        t=node.find('p:spPr/a:xfrm',NS)
                        if t is None:raise ValueError(f'{name}: missing picture transform')
                        matrix=_local_transform(t)*parent
                        bbox=tuple(fitz.Rect(0,0,1,1)*matrix)
                        bl=node.find('p:blipFill/a:blip',NS);media=sr.get(bl.get('{'+NS['r']+'}embed'))
                        cr=node.find('p:blipFill/a:srcRect',NS)
                        crop=tuple(int(cr.get(k,0))/100000 if cr is not None else 0 for k in ('l','t','r','b'))
                        effects=[];reason=None
                        for effect in bl:
                            typ=effect.tag.split('}')[-1]
                            if typ=='extLst':continue  # cached display bitmap already includes edit history
                            if typ=='clrChange':
                                fr=effect.find('a:clrFrom/a:srgbClr',NS);to=effect.find('a:clrTo/a:srgbClr',NS)
                                if fr is None or to is None:raise ValueError(f'{name}: unsupported color representation')
                                alpha=to.find('a:alpha',NS)
                                effects.append({'kind':typ,'from':fr.get('val'),'to':to.get('val'),
                                                'alpha':int(alpha.get('val'))/100000 if alpha is not None else 1})
                            elif typ=='grayscl':effects.append({'kind':typ})
                            elif typ=='biLevel':effects.append({'kind':typ,'threshold':int(effect.get('thresh'))/100000})
                            elif typ=='duotone':
                                first=list(effect)[0]
                                if not ((first.tag.endswith('prstClr') and first.get('val')=='black') or (first.tag.endswith('srgbClr') and first.get('val')=='000000')):
                                    raise ValueError(f'{name}: this converter requires a black-anchored duotone')
                                # The exact output color ramp is measured from this
                                # picture's native PDF. Only colors, never image detail,
                                # are obtained from the exported low-resolution bitmap.
                                effects.append({'kind':typ})
                            else:raise ValueError(f'{name}: unsupported picture effect {typ}')
                        sp=node.find('p:spPr',NS);geom=sp.find('a:prstGeom',NS)
                        if geom is not None and geom.get('prst')!='rect':
                            raise ValueError(f'{name}: nonrectangular clipping is not supported')
                        if any(c.tag.split('}')[-1] in ('effectLst','effectDag','custGeom','scene3d','sp3d') for c in sp):
                            raise ValueError(f'{name}: unsupported picture-shape effect')
                        if not media:raise ValueError(f'{name}: external picture is unavailable')
                        pictures.append(dict(name=name,bbox=bbox,matrix=tuple(matrix),crop=crop,
                                             data=z.read(media),effects=effects,reason=reason,media=media))
            walk(ET.fromstring(z.read(part)).find('p:cSld/p:spTree',NS),fitz.Matrix(1,1))
            slides.append(pictures)
    return slides,size


def _native_export(pptx,pdf):
    if os.name!='nt':raise RuntimeError('Supply base_pdf or run on Windows with desktop PowerPoint.')
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize();app=None;pres=None
    try:
        app=win32com.client.DispatchEx('PowerPoint.Application');app.AutomationSecurity=3
        pres=app.Presentations.Open(str(Path(pptx).resolve()),True,False,False)
        for slide in pres.Slides:slide.SlideShowTransition.Hidden=False
        pres.SaveAs(str(Path(pdf).resolve()),32)
    finally:
        try:
            if pres is not None:pres.Close()
        finally:
            try:
                if app is not None and app.Presentations.Count==0:app.Quit()
            finally:pythoncom.CoUninitialize()


def _luminance(rgb):
    return rgb.convert('L',matrix=(.2126,.7152,.0722,0))


def _apply_effects(pic,doc,native_xref):
    im=Image.open(io.BytesIO(pic['data'])).convert('RGBA')
    rgb=im.convert('RGB');alpha=im.getchannel('A');measurements=[]
    for effect in pic['effects']:
        kind=effect['kind']
        if kind=='clrChange':
            source=tuple(int(effect['from'][i:i+2],16) for i in (0,2,4))
            dest=tuple(int(effect['to'][i:i+2],16) for i in (0,2,4))
            mask=Image.new('L',rgb.size,255)
            for channel,value in zip(rgb.split(),source):
                mask=ImageChops.multiply(mask,channel.point(lambda x,v=value:255 if x==v else 0))
            rgb.paste(dest,mask=mask)
            alpha=Image.composite(alpha.point(lambda x:round(x*effect['alpha'])),alpha,mask)
        elif kind=='grayscl':rgb=Image.merge('RGB',(_luminance(rgb),)*3)
        elif kind=='biLevel':
            threshold=effect['threshold']*255
            bw=_luminance(rgb).point(lambda v:255 if v>=threshold else 0)
            rgb=Image.merge('RGB',(bw,)*3)
        elif kind=='duotone':
            # Match the native Office color ramp, without borrowing its pixels
            # or upscaling it. All image detail is rebuilt from the source bitmap.
            native=fitz.Pixmap(doc,native_xref)
            if native.colorspace.n!=3:native=fitz.Pixmap(fitz.csRGB,native)
            target=Image.open(io.BytesIO(native.tobytes('png'))).convert('RGBA')
            pre=rgb.convert('RGBA');pre.putalpha(alpha)
            small=pre.resize(target.size,Image.Resampling.LANCZOS)
            lum=_luminance(small.convert('RGB'))
            am=ImageChops.darker(small.getchannel('A'),target.getchannel('A')).filter(ImageFilter.MinFilter(3))
            spread=ImageChops.subtract(lum.filter(ImageFilter.MaxFilter(3)),lum.filter(ImageFilter.MinFilter(3)))
            samples=[(v/255,col[:3]) for v,col,a,delta in zip(lum.getdata(),target.getdata(),am.getdata(),spread.getdata())
                     if a>250 and 40<v<250 and delta<15]
            if len(samples)<10:raise ValueError(f'{pic["name"]}: insufficient native color samples for duotone')
            denominator=sum(v*v for v,col in samples)
            ramp=tuple(sum(v*col[k] for v,col in samples)/denominator for k in range(3))
            if max(ramp)>270 or min(ramp)<0:raise ValueError('Invalid measured duotone ramp')
            rmse=(sum((v*ramp[k]-col[k])**2 for v,col in samples for k in range(3))/(3*len(samples)))**.5
            if rmse>5:raise ValueError(f'{pic["name"]}: duotone color verification failed ({rmse:.2f})')
            gray=_luminance(rgb)
            rgb=Image.merge('RGB',tuple(gray.point(lambda v,k=k:round(max(0,min(255,v*ramp[k]/255)))) for k in range(3)))
            measurements.append({'effect':'duotone','color_rmse':round(rmse,4),'samples':len(samples)})
        else:raise ValueError(f'Unsupported effect: {kind}')
    return rgb,alpha,measurements


def _embed(doc,pic,native_xref,cache):
    rgb,alpha,measurements=_apply_effects(pic,doc,native_xref)
    raw=rgb.tobytes();araw=alpha.tobytes()
    rhash=hashlib.sha256(raw).hexdigest();ahash=hashlib.sha256(araw).hexdigest()
    key=(rgb.size,rhash,ahash)
    if key in cache:return cache[key]
    maskref=None
    if alpha.getextrema()!=(255,255):
        maskref=doc.get_new_xref()
        doc.update_object(maskref,f'<< /Type /XObject /Subtype /Image /Width {rgb.width} /Height {rgb.height} /ColorSpace /DeviceGray /BitsPerComponent 8 >>')
        doc.update_stream(maskref,araw,compress=True)
    xref=doc.get_new_xref();smask=f'/SMask {maskref} 0 R' if maskref else ''
    doc.update_object(xref,f'<< /Type /XObject /Subtype /Image /Width {rgb.width} /Height {rgb.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 {smask} >>')
    doc.update_stream(xref,raw,compress=True)
    record={'xref':xref,'size':rgb.size,'rgb_sha256':rhash,'alpha_sha256':ahash if maskref else None,'effects':measurements}
    cache[key]=record;return record


def _geometry_options(pic):
    yield pic
    l,t,r,b=pic['crop']
    if min(l,t,r,b)>=0:return
    fx=1-l-r;fy=1-t-b
    if fx<=0 or fy<=0:return
    rect=fitz.Rect(max(0,-l/fx),max(0,-t/fy),min(1,(1-l)/fx),min(1,(1-t)/fy))
    visible=tuple(rect*fitz.Matrix(*pic['matrix']))
    yield dict(pic,bbox=visible,crop=tuple(max(0,v) for v in pic['crop']))


def _match_pictures(pics, infos, page_no):
    """Minimum-error ordered matching; extra PDF images may be skipped."""
    supported=[p for p in pics if not p['reason']]
    n,m=len(supported),len(infos)
    inf=float('inf')
    costs=[[inf]*(m+1) for _ in range(n+1)]
    ways=[[0]*(m+1) for _ in range(n+1)]
    take=[[False]*m for _ in range(n)]
    chosen={}
    for j in range(m+1):costs[n][j]=0;ways[n][j]=1
    for i in range(n-1,-1,-1):
        for j in range(m-1,-1,-1):
            best=costs[i][j+1];count=ways[i][j+1]
            variants=list(_geometry_options(supported[i]))
            errors=[max(abs(a-b) for a,b in zip(v['bbox'],infos[j]['bbox'])) for v in variants]
            k=min(range(len(errors)),key=errors.__getitem__)
            err=errors[k]
            chosen[i,j]=variants[k]
            match=err+costs[i+1][j+1] if err<=.2 and infos[j]['xref'] else inf
            if match<best-1e-7:
                best=match;count=ways[i+1][j+1];take[i][j]=True
            elif match!=inf and abs(match-best)<=1e-7:
                count=min(2,count+ways[i+1][j+1])
            costs[i][j]=best;ways[i][j]=count
    if costs[0][0]==inf:
        raise ValueError(f'Page {page_no}: cannot locate all {n} supported pictures among {m} PDF images by position. No PDF written.')
    if ways[0][0]>1:
        raise ValueError(f'Page {page_no}: multiple equally good picture mappings. No PDF written; inspect overlapping image objects.')
    matches=[];i=j=0
    while i<n:
        if take[i][j]:matches.append((chosen[i,j],j));i+=1
        j+=1
    return matches


def export(ppt_path,pdf_path,base_pdf=None):
    source=Path(ppt_path).resolve();output=Path(pdf_path).resolve()
    if output==source or (base_pdf and output==Path(base_pdf).resolve()):
        raise ValueError('Output must differ from both input files.')
    pictures,size=_read(source);output.parent.mkdir(parents=True,exist_ok=True)
    report={'restored':[],'retained_effects':[],'native_images_retained':[]}
    with tempfile.TemporaryDirectory(prefix='ppt_full_resolution_') as td:
        base=Path(base_pdf).resolve() if base_pdf else Path(td)/'native.pdf'
        if base_pdf is None:_native_export(source,base)
        with fitz.open(base) as doc:
            if len(doc)!=len(pictures):raise ValueError('Slide/page count mismatch: export all slides, one per page.')
            texts=[p.get_text() for p in doc]
            native_streams=[[hashlib.sha256(doc.xref_stream(x)).hexdigest() for x in p.get_contents()] for p in doc]
            plans={};skipped_refs=set();page_matches=[];original_counts=[]
            for page_no,(page,pics) in enumerate(zip(doc,pictures),1):
                if max(abs(page.rect.width-size[0]),abs(page.rect.height-size[1]))>.1:
                    raise ValueError(f'Page {page_no}: slide dimensions do not match.')
                infos=page.get_image_info(xrefs=True)
                matches=_match_pictures(pics,infos,page_no)
                page_matches.append(matches);original_counts.append(len(infos))
                matched_indexes={j for _,j in matches}
                for pic in pics:
                    if pic['reason']:
                        report['retained_effects'].append({'page':page_no,'picture':pic['name'],'reason':pic['reason']})
                for j,info in enumerate(infos):
                    if j not in matched_indexes:
                        skipped_refs.add(info['xref'])
                        report['native_images_retained'].append({'page':page_no,'image_index':j,'pixels':[info['width'],info['height']]})
                for pic,j in matches:
                    info=infos[j];old=info['xref']
                    sig=(hashlib.sha256(pic['data']).hexdigest(),pic['crop'],json.dumps(pic['effects'],sort_keys=True))
                    if old in plans and plans[old]['sig']!=sig:
                        raise ValueError('A shared PDF image has conflicting source crops. No PDF written.')
                    plans[old]={'sig':sig,'pic':pic}
                    report['restored'].append({'page':page_no,'picture':pic['name'],
                          'exported_pixels':[info['width'],info['height']],
                          'original_pixels':list(Image.open(io.BytesIO(pic['data'])).size)})
            if skipped_refs.intersection(plans):raise ValueError('A PDF image is shared by edited and unedited pictures; cannot safely replace it.')
            if not plans:raise ValueError('No images eligible for direct restoration. No PDF written.')
            cache={};expected_forms={}
            for old,plan in plans.items():
                pic=plan['pic'];embedded=_embed(doc,pic,old,cache);ref=embedded['xref'];dimensions=embedded['size']
                l,t,r,b=pic['crop'];fx=1-l-r;fy=1-t-b
                if fx<=0 or fy<=0:raise ValueError('Invalid picture crop.')
                # A PDF image occupies a unit square. Substitute a clipped Form
                # in that same square, containing the complete original bitmap.
                # No resampling or integer rounding of crop boundaries occurs.
                doc.update_object(old,f'<< /Type /XObject /Subtype /Form /FormType 1 /RestorationId (restored_{old}) /BBox [0 0 1 1] /Resources << /XObject << /Original {ref} 0 R >> >> >>')
                stream=f'q 0 0 1 1 re W n {1/fx:.12f} 0 0 {1/fy:.12f} {-l/fx:.12f} {-b/fy:.12f} cm /Original Do Q'.encode('ascii')
                doc.update_stream(old,stream,compress=True)
                expected_forms[f'restored_{old}']={'image':embedded,'stream':stream,'name':pic['name']}
            # Original page content streams, text, and vector paths were not edited.
            temp=output.with_name(output.stem+'.writing.pdf')
            try:
                doc.save(temp,garbage=4,deflate=True)
                with fitz.open(temp) as check:
                    if [p.get_text() for p in check]!=texts:raise RuntimeError('Native text verification failed.')
                    if [[hashlib.sha256(check.xref_stream(x)).hexdigest() for x in p.get_contents()] for p in check]!=native_streams:
                        raise RuntimeError('Native page drawing instructions changed')
                    # Verify tagged PDF objects directly. get_image_info() order
                    # and caching vary across PyMuPDF versions and nested Forms.
                    found={}
                    for xr in range(1,check.xref_length()):
                        kind,tag=check.xref_get_key(xr,'RestorationId')
                        if kind=='string' and tag in expected_forms:found[tag]=xr
                    if set(found)!=set(expected_forms):raise RuntimeError('Missing restored PDF image objects')
                    verified_images=set()
                    for tag,expected in expected_forms.items():
                        form=found[tag]
                        if check.xref_stream(form)!=expected['stream']:
                            raise RuntimeError(f'Crop verification failed: {expected["name"]}')
                        kind,value=check.xref_get_key(form,'Resources/XObject/Original')
                        if kind!='xref':raise RuntimeError('Restored image reference is missing')
                        image=int(value.split()[0]);record=expected['image']
                        dims=tuple(int(check.xref_get_key(image,k)[1]) for k in ('Width','Height'))
                        if dims!=tuple(record['size']):raise RuntimeError(f'Image dimensions changed: {expected["name"]}')
                        if image not in verified_images:
                            if hashlib.sha256(check.xref_stream(image)).hexdigest()!=record['rgb_sha256']:
                                raise RuntimeError(f'Image pixels changed: {expected["name"]}')
                            kind,mask=check.xref_get_key(image,'SMask')
                            if record['alpha_sha256']:
                                if kind!='xref' or hashlib.sha256(check.xref_stream(int(mask.split()[0]))).hexdigest()!=record['alpha_sha256']:
                                    raise RuntimeError(f'Image transparency changed: {expected["name"]}')
                            verified_images.add(image)
                    report['verification']={'pages':len(check),'picture_placements':len(report['restored']),
                        'restored_objects':len(found),'unique_full_resolution_images':len(verified_images),
                        'pixel_hashes_verified':True,'crop_streams_verified':True,'native_text_verified':True}

                temp.replace(output)
            finally:temp.unlink(missing_ok=True)
    output.with_suffix('.restoration.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
    print(f'Saved {output}: {len(report["restored"])} picture placements restored at source resolution.')
    for item in report['retained_effects']:
        print(f'UNCHANGED: page {item["page"]}, {item["picture"]}: {item["reason"]}')
    return report


if __name__=='__main__':
    ap=argparse.ArgumentParser(description=__doc__,formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('pptx');ap.add_argument('pdf');ap.add_argument('--base-pdf')
    args=ap.parse_args();export(args.pptx,args.pdf,base_pdf=args.base_pdf)
    