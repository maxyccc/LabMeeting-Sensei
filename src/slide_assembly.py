from io import BytesIO
from typing import Union
from pptx import Presentation
from pptx.util import Inches
from PIL import Image
import os

from src.basic_layouts import Layout
from src.settings import DPI, SLIDE_WIDTH, SLIDE_HEIGHT
from src.images import Size, paste_offset


def create_presentation(
        width: float = SLIDE_WIDTH,
        height: float = SLIDE_HEIGHT,
):
    prs = Presentation()
    prs.slide_width = Inches(width)
    prs.slide_height = Inches(height)
    return prs


def insert_slide(content: Union[str, Layout], prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 6 is the blank layout
    if isinstance(content, str):
        with Image.open(content) as im:
            im_w, im_h = im.size

        slide_size = Size(prs.slide_width, prs.slide_height)
        scale = min(slide_size.w / im_w, slide_size.h / im_h)
        pic_size = Size(int(im_w * scale), int(im_h * scale))
        left, top = paste_offset(slide_size, pic_size)
        slide.shapes.add_picture(content, left, top, width=pic_size.w, height=pic_size.h)
        return slide

    slide_w_px = int((prs.slide_width / Inches(1)) * DPI)
    slide_h_px = int((prs.slide_height / Inches(1)) * DPI)
    placements = content.place(Size(slide_w_px, slide_h_px))
    for x, y, img in placements:
        buf = BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        slide.shapes.add_picture(
            buf,
            Inches(x / DPI),
            Inches(y / DPI),
            width=Inches(img.width / DPI),
            height=Inches(img.height / DPI),
        )
    return slide


def save_presentation(prs: Presentation, save_path: str):
    try:
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        prs.save(save_path)
        print(f"\n✅ Successfully created presentation: {save_path}")
    except Exception as e:
        print(f"\n❌ Failed to save presentation. Error: {e}")
