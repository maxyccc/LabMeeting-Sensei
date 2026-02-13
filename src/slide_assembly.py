
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image
import os

from src.settings import SLIDE_WIDTH, SLIDE_HEIGHT
from src.images import Size, paste_offset


def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH)
    prs.slide_height = Inches(SLIDE_HEIGHT)
    return prs


def insert_slide(im_path: str, prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 6 is the blank layout
    with Image.open(im_path) as im:
        im_w, im_h = im.size

    slide_size = Size(prs.slide_width, prs.slide_height)
    scale = min(slide_size.w / im_w, slide_size.h / im_h)
    pic_size = Size(int(im_w * scale), int(im_h * scale))
    left, top = paste_offset(slide_size, pic_size)
    slide.shapes.add_picture(im_path, left, top, width=pic_size.w, height=pic_size.h)
    return slide


def save_presentation(prs: Presentation, save_path: str):
    try:
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        prs.save(save_path)
        print(f"\n✅ Successfully created presentation: {save_path}")
    except Exception as e:
        print(f"\n❌ Failed to save presentation. Error: {e}")
