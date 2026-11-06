"""Restore PPTX source images in an existing PDF while preserving its layout.

Match pictures by position, apply supported color effects at source resolution,
replace PDF image objects with cropped Forms, then verify the saved result.
The caller must supply base_pdf; this module does not export PowerPoint to PDF.
"""

import argparse
import hashlib
import io
import json
import posixpath
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from PIL import Image, ImageChops, ImageFilter
import pymupdf as fitz


NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


def _read_relationships(archive, part):
    folder, name = posixpath.split(part)
    path = posixpath.join(folder, '_rels', name + '.rels')
    return {
        relationship.get('Id'): posixpath.normpath(
            posixpath.join(folder, relationship.get('Target'))
        )
        for relationship in ET.fromstring(archive.read(path))
        if relationship.get('TargetMode') != 'External'
    }


def _local_transform(transform, group=False):
    """Map a picture's unit square or a group's child coordinates to its parent."""
    offset = transform.find('a:off', NS)
    extent = transform.find('a:ext', NS)
    # DrawingML uses 12,700 EMUs per PDF point.
    x, y = [int(offset.get(key)) / 12700 for key in ('x', 'y')]
    width, height = [int(extent.get(key)) / 12700 for key in ('cx', 'cy')]
    if group:
        child_offset = transform.find('a:chOff', NS)
        child_extent = transform.find('a:chExt', NS)
        child_width, child_height = [
            int(child_extent.get(key)) / 12700 for key in ('cx', 'cy')
        ]
        child_x, child_y = [
            int(child_offset.get(key)) / 12700 for key in ('x', 'y')
        ]
        scale_x = width / child_width if child_width else 1
        scale_y = height / child_height if child_height else 1
        matrix = fitz.Matrix(
            scale_x, 0, 0, scale_y, x - scale_x * child_x, y - scale_y * child_y
        )
    else:
        matrix = fitz.Matrix(width, 0, 0, height, x, y)

    flip_x = -1 if transform.get('flipH') in ('1', 'true') else 1
    flip_y = -1 if transform.get('flipV') in ('1', 'true') else 1
    angle = int(transform.get('rot', 0)) / 60000
    center_x, center_y = x + width / 2, y + height / 2
    # Flip and rotate about the shape's center before composing with its parent.
    around_center = (
        fitz.Matrix(1, 0, 0, 1, -center_x, -center_y)
        * fitz.Matrix(flip_x, flip_y)
        * fitz.Matrix(angle)
        * fitz.Matrix(1, 0, 0, 1, center_x, center_y)
    )
    return matrix * around_center


def _read_presentation(pptx):
    """Read source images, crops, effects, and slide-space geometry in draw order."""
    slides = []
    with zipfile.ZipFile(pptx) as archive:
        presentation = ET.fromstring(archive.read('ppt/presentation.xml'))
        relationships = _read_relationships(archive, 'ppt/presentation.xml')
        slide_size = presentation.find('p:sldSz', NS)
        size = tuple(int(slide_size.get(key)) / 12700 for key in ('cx', 'cy'))
        for slide_id in presentation.findall('p:sldIdLst/p:sldId', NS):
            part = relationships[slide_id.get('{' + NS['r'] + '}id')]
            slide_relationships = _read_relationships(archive, part)
            pictures = []

            def walk(tree, parent_matrix):
                for node in tree:
                    tag = node.tag.split('}')[-1]
                    if tag == 'grpSp':
                        if not node.findall('.//p:pic', NS):
                            continue
                        transform = node.find('p:grpSpPr/a:xfrm', NS)
                        walk(node, _local_transform(transform, True) * parent_matrix)
                        continue
                    if tag != 'pic':
                        continue
                    name = node.find('p:nvPicPr/p:cNvPr', NS).get('name', 'Picture')
                    transform = node.find('p:spPr/a:xfrm', NS)
                    if transform is None:
                        raise ValueError(f'{name}: missing picture transform')
                    matrix = _local_transform(transform) * parent_matrix
                    bbox = tuple(fitz.Rect(0, 0, 1, 1) * matrix)
                    blip = node.find('p:blipFill/a:blip', NS)
                    media = slide_relationships.get(blip.get('{' + NS['r'] + '}embed'))
                    crop_rect = node.find('p:blipFill/a:srcRect', NS)
                    crop = tuple(
                        int(crop_rect.get(edge, 0)) / 100000
                        if crop_rect is not None else 0
                        for edge in ('l', 't', 'r', 'b')
                    )
                    effects = []
                    for effect in blip:
                        kind = effect.tag.split('}')[-1]
                        if kind == 'extLst':
                            # The cached display bitmap already includes edit history.
                            continue
                        if kind == 'clrChange':
                            source_color = effect.find('a:clrFrom/a:srgbClr', NS)
                            target_color = effect.find('a:clrTo/a:srgbClr', NS)
                            if source_color is None or target_color is None:
                                raise ValueError(f'{name}: unsupported color representation')
                            alpha = target_color.find('a:alpha', NS)
                            effects.append({
                                'kind': kind,
                                'from': source_color.get('val'),
                                'to': target_color.get('val'),
                                'alpha': int(alpha.get('val')) / 100000
                                if alpha is not None else 1,
                            })
                        elif kind == 'grayscl':
                            effects.append({'kind': kind})
                        elif kind == 'biLevel':
                            effects.append({
                                'kind': kind,
                                'threshold': int(effect.get('thresh')) / 100000,
                            })
                        elif kind == 'duotone':
                            first_color = list(effect)[0]
                            preset_black = (
                                first_color.tag.endswith('prstClr')
                                and first_color.get('val') == 'black'
                            )
                            rgb_black = (
                                first_color.tag.endswith('srgbClr')
                                and first_color.get('val') == '000000'
                            )
                            if not (preset_black or rgb_black):
                                raise ValueError(
                                    f'{name}: this converter requires a black-anchored duotone'
                                )
                            effects.append({'kind': kind})
                        else:
                            raise ValueError(f'{name}: unsupported picture effect {kind}')

                    shape_properties = node.find('p:spPr', NS)
                    geometry = shape_properties.find('a:prstGeom', NS)
                    if geometry is not None and geometry.get('prst') != 'rect':
                        raise ValueError(f'{name}: nonrectangular clipping is not supported')
                    if any(
                        child.tag.split('}')[-1] in (
                            'effectLst', 'effectDag', 'custGeom', 'scene3d', 'sp3d'
                        )
                        for child in shape_properties
                    ):
                        raise ValueError(f'{name}: unsupported picture-shape effect')
                    if not media:
                        raise ValueError(f'{name}: external picture is unavailable')
                    pictures.append({
                        'name': name, 'bbox': bbox, 'matrix': tuple(matrix),
                        'crop': crop, 'data': archive.read(media),
                        'effects': effects, 'reason': None, 'media': media,
                    })

            slide_tree = ET.fromstring(archive.read(part)).find('p:cSld/p:spTree', NS)
            walk(slide_tree, fitz.Matrix(1, 1))
            slides.append(pictures)
    return slides, size


def _luminance(rgb):
    return rgb.convert('L', matrix=(.2126, .7152, .0722, 0))


def _apply_effects(picture, document, native_xref):
    """Apply color effects to source pixels, measuring duotone colors from the PDF."""
    image = Image.open(io.BytesIO(picture['data'])).convert('RGBA')
    rgb = image.convert('RGB')
    alpha = image.getchannel('A')
    measurements = []
    for effect in picture['effects']:
        kind = effect['kind']
        if kind == 'clrChange':
            source_color = tuple(int(effect['from'][i:i + 2], 16) for i in (0, 2, 4))
            target_color = tuple(int(effect['to'][i:i + 2], 16) for i in (0, 2, 4))
            mask = Image.new('L', rgb.size, 255)
            for channel, value in zip(rgb.split(), source_color):
                matching_pixels = channel.point(
                    lambda pixel, value=value: 255 if pixel == value else 0
                )
                mask = ImageChops.multiply(mask, matching_pixels)
            rgb.paste(target_color, mask=mask)
            alpha = Image.composite(
                alpha.point(lambda pixel: round(pixel * effect['alpha'])), alpha, mask
            )
        elif kind == 'grayscl':
            rgb = Image.merge('RGB', (_luminance(rgb),) * 3)
        elif kind == 'biLevel':
            threshold = effect['threshold'] * 255
            black_and_white = _luminance(rgb).point(lambda value: 255 if value >= threshold else 0)
            rgb = Image.merge('RGB', (black_and_white,) * 3)
        elif kind == 'duotone':
            native = fitz.Pixmap(document, native_xref)
            if native.colorspace.n != 3:
                native = fitz.Pixmap(fitz.csRGB, native)
            target = Image.open(io.BytesIO(native.tobytes('png'))).convert('RGBA')
            source_rgba = rgb.convert('RGBA')
            source_rgba.putalpha(alpha)
            resized_source = source_rgba.resize(target.size, Image.Resampling.LANCZOS)
            luminance = _luminance(resized_source.convert('RGB'))
            opaque_mask = ImageChops.darker(
                resized_source.getchannel('A'), target.getchannel('A')
            ).filter(ImageFilter.MinFilter(3))
            local_spread = ImageChops.subtract(
                luminance.filter(ImageFilter.MaxFilter(3)),
                luminance.filter(ImageFilter.MinFilter(3)),
            )
            # Use opaque, locally flat midtones to fit a least-squares color ramp
            # through black. Only colors come from the PDF, never image detail.
            samples = [
                (value / 255, color[:3])
                for value, color, opacity, spread in zip(
                    luminance.getdata(), target.getdata(),
                    opaque_mask.getdata(), local_spread.getdata(),
                )
                if opacity > 250 and 40 < value < 250 and spread < 15
            ]
            if len(samples) < 10:
                raise ValueError(f'{picture["name"]}: insufficient native color samples for duotone')
            denominator = sum(value * value for value, color in samples)
            ramp = tuple(
                sum(value * color[channel] for value, color in samples) / denominator
                for channel in range(3)
            )
            if max(ramp) > 270 or min(ramp) < 0:
                raise ValueError('Invalid measured duotone ramp')
            squared_error = sum(
                (value * ramp[channel] - color[channel]) ** 2
                for value, color in samples for channel in range(3)
            )
            rmse = (squared_error / (3 * len(samples))) ** .5
            if rmse > 5:
                raise ValueError(f'{picture["name"]}: duotone color verification failed ({rmse:.2f})')
            gray = _luminance(rgb)
            channels = tuple(
                gray.point(lambda value, channel=channel: round(
                    max(0, min(255, value * ramp[channel] / 255))
                ))
                for channel in range(3)
            )
            rgb = Image.merge('RGB', channels)
            measurements.append({
                'effect': 'duotone', 'color_rmse': round(rmse, 4), 'samples': len(samples),
            })
        else:
            raise ValueError(f'Unsupported effect: {kind}')
    return rgb, alpha, measurements


def _embed_image(document, picture, native_xref, cache):
    rgb, alpha, measurements = _apply_effects(picture, document, native_xref)
    rgb_bytes = rgb.tobytes()
    alpha_bytes = alpha.tobytes()
    rgb_hash = hashlib.sha256(rgb_bytes).hexdigest()
    alpha_hash = hashlib.sha256(alpha_bytes).hexdigest()
    cache_key = (rgb.size, rgb_hash, alpha_hash)
    if cache_key in cache:
        return cache[cache_key]

    mask_xref = None
    if alpha.getextrema() != (255, 255):
        mask_xref = document.get_new_xref()
        document.update_object(mask_xref, (
            f'<< /Type /XObject /Subtype /Image /Width {rgb.width} /Height {rgb.height} '
            '/ColorSpace /DeviceGray /BitsPerComponent 8 >>'
        ))
        document.update_stream(mask_xref, alpha_bytes, compress=True)
    image_xref = document.get_new_xref()
    soft_mask = f'/SMask {mask_xref} 0 R' if mask_xref else ''
    document.update_object(image_xref, (
        f'<< /Type /XObject /Subtype /Image /Width {rgb.width} /Height {rgb.height} '
        f'/ColorSpace /DeviceRGB /BitsPerComponent 8 {soft_mask} >>'
    ))
    document.update_stream(image_xref, rgb_bytes, compress=True)
    record = {
        'xref': image_xref, 'size': rgb.size, 'rgb_sha256': rgb_hash,
        'alpha_sha256': alpha_hash if mask_xref else None, 'effects': measurements,
    }
    cache[cache_key] = record
    return record


def _geometry_options(picture):
    """Also allow the PDF's visible-image bounds for negative (outset) crops."""
    yield picture
    left, top, right, bottom = picture['crop']
    if min(left, top, right, bottom) >= 0:
        return
    crop_width = 1 - left - right
    crop_height = 1 - top - bottom
    if crop_width <= 0 or crop_height <= 0:
        return
    # Negative crops add space outside the bitmap. Intersect its bounds with the
    # picture's unit square, then transform that visible region into slide space.
    visible_rect = fitz.Rect(
        max(0, -left / crop_width), max(0, -top / crop_height),
        min(1, (1 - left) / crop_width), min(1, (1 - top) / crop_height),
    )
    visible_bbox = tuple(visible_rect * fitz.Matrix(*picture['matrix']))
    yield dict(picture, bbox=visible_bbox, crop=tuple(max(0, edge) for edge in picture['crop']))


def _match_pictures(pictures, image_infos, page_no):
    """Find a unique minimum-error ordered match, allowing extra PDF images."""
    supported = [picture for picture in pictures if not picture['reason']]
    picture_count, image_count = len(supported), len(image_infos)
    infinity = float('inf')
    costs = [[infinity] * (image_count + 1) for _ in range(picture_count + 1)]
    ways = [[0] * (image_count + 1) for _ in range(picture_count + 1)]
    take_match = [[False] * image_count for _ in range(picture_count)]
    chosen_variants = {}
    # Each cell matches the remaining suffixes: skip this PDF image, or pair it
    # with this picture. Count optimal paths up to two to detect ambiguity.
    for image_index in range(image_count + 1):
        costs[picture_count][image_index] = 0
        ways[picture_count][image_index] = 1
    for picture_index in range(picture_count - 1, -1, -1):
        for image_index in range(image_count - 1, -1, -1):
            best = costs[picture_index][image_index + 1]
            count = ways[picture_index][image_index + 1]
            variants = list(_geometry_options(supported[picture_index]))
            errors = [
                max(abs(a - b) for a, b in zip(variant['bbox'], image_infos[image_index]['bbox']))
                for variant in variants
            ]
            best_variant = min(range(len(errors)), key=errors.__getitem__)
            error = errors[best_variant]
            chosen_variants[picture_index, image_index] = variants[best_variant]
            match_cost = (
                error + costs[picture_index + 1][image_index + 1]
                if error <= .2 and image_infos[image_index]['xref'] else infinity
            )
            if match_cost < best - 1e-7:
                best = match_cost
                count = ways[picture_index + 1][image_index + 1]
                take_match[picture_index][image_index] = True
            elif match_cost != infinity and abs(match_cost - best) <= 1e-7:
                count = min(2, count + ways[picture_index + 1][image_index + 1])
            costs[picture_index][image_index] = best
            ways[picture_index][image_index] = count
    if costs[0][0] == infinity:
        raise ValueError(
            f'Page {page_no}: cannot locate all {picture_count} supported pictures '
            f'among {image_count} PDF images by position. No PDF written.'
        )
    if ways[0][0] > 1:
        raise ValueError(
            f'Page {page_no}: multiple equally good picture mappings. '
            'No PDF written; inspect overlapping image objects.'
        )
    matches = []
    picture_index = image_index = 0
    while picture_index < picture_count:
        if take_match[picture_index][image_index]:
            matches.append((chosen_variants[picture_index, image_index], image_index))
            picture_index += 1
        image_index += 1
    return matches


def _plan_replacements(document, slides, slide_size, report):
    """Match placements and reject conflicting uses of shared PDF image objects."""
    plans = {}
    skipped_refs = set()
    for page_no, (page, pictures) in enumerate(zip(document, slides), 1):
        if max(abs(page.rect.width - slide_size[0]), abs(page.rect.height - slide_size[1])) > .1:
            raise ValueError(f'Page {page_no}: slide dimensions do not match.')
        image_infos = page.get_image_info(xrefs=True)
        matches = _match_pictures(pictures, image_infos, page_no)
        matched_indexes = {image_index for _, image_index in matches}
        for picture in pictures:
            if picture['reason']:
                report['retained_effects'].append({
                    'page': page_no, 'picture': picture['name'], 'reason': picture['reason'],
                })
        for image_index, info in enumerate(image_infos):
            if image_index not in matched_indexes:
                skipped_refs.add(info['xref'])
                report['native_images_retained'].append({
                    'page': page_no, 'image_index': image_index,
                    'pixels': [info['width'], info['height']],
                })
        for picture, image_index in matches:
            info = image_infos[image_index]
            native_xref = info['xref']
            signature = (
                hashlib.sha256(picture['data']).hexdigest(), picture['crop'],
                json.dumps(picture['effects'], sort_keys=True),
            )
            if native_xref in plans and plans[native_xref]['signature'] != signature:
                raise ValueError('A shared PDF image has conflicting source crops. No PDF written.')
            plans[native_xref] = {'signature': signature, 'picture': picture}
            report['restored'].append({
                'page': page_no, 'picture': picture['name'],
                'exported_pixels': [info['width'], info['height']],
                'original_pixels': list(Image.open(io.BytesIO(picture['data'])).size),
            })
    if skipped_refs.intersection(plans):
        raise ValueError('A PDF image is shared by edited and unedited pictures; cannot safely replace it.')
    if not plans:
        raise ValueError('No images eligible for direct restoration. No PDF written.')
    return plans


def _replace_images(document, plans):
    """Replace image XObjects in place with Forms that draw the source bitmap."""
    cache = {}
    expected_forms = {}
    for native_xref, plan in plans.items():
        picture = plan['picture']
        embedded = _embed_image(document, picture, native_xref, cache)
        image_xref = embedded['xref']
        left, top, right, bottom = picture['crop']
        crop_width = 1 - left - right
        crop_height = 1 - top - bottom
        if crop_width <= 0 or crop_height <= 0:
            raise ValueError('Invalid picture crop.')
        # Keep the old xref so every existing placement still refers to this
        # object. A Form occupies the same unit square, clips it, and draws the
        # full bitmap inside; page content and placement matrices stay intact.
        document.update_object(native_xref, (
            f'<< /Type /XObject /Subtype /Form /FormType 1 '
            f'/RestorationId (restored_{native_xref}) /BBox [0 0 1 1] '
            f'/Resources << /XObject << /Original {image_xref} 0 R >> >> >>'
        ))
        # Invert the fractional crop without resampling or rounding pixels.
        # PDF's vertical axis starts at the bottom, so translation uses bottom.
        stream = (
            f'q 0 0 1 1 re W n {1 / crop_width:.12f} 0 0 {1 / crop_height:.12f} '
            f'{-left / crop_width:.12f} {-bottom / crop_height:.12f} cm /Original Do Q'
        ).encode('ascii')
        document.update_stream(native_xref, stream, compress=True)
        expected_forms[f'restored_{native_xref}'] = {
            'image': embedded, 'stream': stream, 'name': picture['name'],
        }
    return expected_forms


def _verify_pdf(document, texts, native_streams, expected_forms, placement_count):
    """Verify native content, crop streams, and exact restored RGB/alpha bytes."""
    if [page.get_text() for page in document] != texts:
        raise RuntimeError('Native text verification failed.')
    saved_streams = [
        [hashlib.sha256(document.xref_stream(xref)).hexdigest() for xref in page.get_contents()]
        for page in document
    ]
    if saved_streams != native_streams:
        raise RuntimeError('Native page drawing instructions changed')
    # Saving can renumber objects. Find tagged Forms directly: image enumeration
    # order and caching vary across PyMuPDF versions and nested Forms.
    found = {}
    for xref in range(1, document.xref_length()):
        kind, tag = document.xref_get_key(xref, 'RestorationId')
        if kind == 'string' and tag in expected_forms:
            found[tag] = xref
    if set(found) != set(expected_forms):
        raise RuntimeError('Missing restored PDF image objects')
    verified_images = set()
    for tag, expected in expected_forms.items():
        form_xref = found[tag]
        if document.xref_stream(form_xref) != expected['stream']:
            raise RuntimeError(f'Crop verification failed: {expected["name"]}')
        kind, value = document.xref_get_key(form_xref, 'Resources/XObject/Original')
        if kind != 'xref':
            raise RuntimeError('Restored image reference is missing')
        image_xref = int(value.split()[0])
        record = expected['image']
        dimensions = tuple(
            int(document.xref_get_key(image_xref, key)[1]) for key in ('Width', 'Height')
        )
        if dimensions != tuple(record['size']):
            raise RuntimeError(f'Image dimensions changed: {expected["name"]}')
        if image_xref not in verified_images:
            if hashlib.sha256(document.xref_stream(image_xref)).hexdigest() != record['rgb_sha256']:
                raise RuntimeError(f'Image pixels changed: {expected["name"]}')
            kind, mask = document.xref_get_key(image_xref, 'SMask')
            if record['alpha_sha256']:
                alpha_matches = kind == 'xref' and hashlib.sha256(
                    document.xref_stream(int(mask.split()[0]))
                ).hexdigest() == record['alpha_sha256']
                if not alpha_matches:
                    raise RuntimeError(f'Image transparency changed: {expected["name"]}')
            verified_images.add(image_xref)
    return {
        'pages': len(document), 'picture_placements': placement_count,
        'restored_objects': len(found), 'unique_full_resolution_images': len(verified_images),
        'pixel_hashes_verified': True, 'crop_streams_verified': True, 'native_text_verified': True,
    }


def export(ppt_path, pdf_path, base_pdf):
    """Restore PPTX images into base_pdf and write a verified PDF and JSON report."""
    source = Path(ppt_path).resolve()
    output = Path(pdf_path).resolve()
    if output == source or (base_pdf and output == Path(base_pdf).resolve()):
        raise ValueError('Output must differ from both input files.')
    slides, slide_size = _read_presentation(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    report = {'restored': [], 'retained_effects': [], 'native_images_retained': []}
    with fitz.open(Path(base_pdf).resolve()) as document:
        if len(document) != len(slides):
            raise ValueError('Slide/page count mismatch: export all slides, one per page.')
        texts = [page.get_text() for page in document]
        native_streams = [
            [hashlib.sha256(document.xref_stream(xref)).hexdigest() for xref in page.get_contents()]
            for page in document
        ]
        plans = _plan_replacements(document, slides, slide_size, report)
        expected_forms = _replace_images(document, plans)

        temporary_pdf = output.with_name(output.stem + '.writing.pdf')
        try:
            document.save(temporary_pdf, garbage=4, deflate=True)
            with fitz.open(temporary_pdf) as saved_document:
                report['verification'] = _verify_pdf(
                    saved_document, texts, native_streams, expected_forms,
                    len(report['restored']),
                )
            temporary_pdf.replace(output)
        finally:
            temporary_pdf.unlink(missing_ok=True)

    output.with_suffix('.restoration.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
    print(f'Saved {output}: {len(report["restored"])} picture placements restored at source resolution.')
    for item in report['retained_effects']:
        print(f'UNCHANGED: page {item["page"]}, {item["picture"]}: {item["reason"]}')
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument('pptx', required=True)
    parser.add_argument('pdf', required=True)
    parser.add_argument('--base-pdf', required=True)
    args = parser.parse_args()
    export(args.pptx, args.pdf, base_pdf=args.base_pdf)
