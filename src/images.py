from dataclasses import dataclass
from enum import Enum
from typing import Tuple, Union

from PIL import Image, ImageOps

from src.settings import BG_COLOR, RESAMPLE


@dataclass(frozen=True)
class Padding:
    left: int
    top: int
    right: int
    bottom: int


def formalize_padding(pad: Union[int, Tuple[int, int], Tuple[int, int, int, int], Padding]) -> Padding:
    if isinstance(pad, Padding):
        return pad
    if isinstance(pad, int):
        return Padding(pad, pad, pad, pad)
    if isinstance(pad, tuple):
        if len(pad) == 2:
            return Padding(pad[0], pad[1], pad[0], pad[1])
        if len(pad) == 4:
            return Padding(*pad)
    raise ValueError(f"Invalid padding: {pad}")


@dataclass(frozen=True)
class Size:
    w: int
    h: int

    def inset(self, pad: Union[int, Tuple[int, int], Tuple[int, int, int, int], Padding]) -> "Size":
        pad = formalize_padding(pad)
        return Size(max(1, self.w - pad.left - pad.right), max(1, self.h - pad.top - pad.bottom))
    
    def expand(self, pad: Union[int, Tuple[int, int], Tuple[int, int, int, int], Padding]) -> "Size":
        pad = formalize_padding(pad)
        return Size(self.w + pad.left + pad.right, self.h + pad.top + pad.bottom)


class FitMode(str, Enum):
    CONTAIN = "contain"   # keep aspect, letterbox
    STRETCH = "stretch"   # distort
    COVER = "cover"       # keep aspect, crop to fill
    NONE = "none"         # no scaling (center)


def load_image(path: str) -> Image.Image:
    im = Image.open(path)
    im = im.convert("RGBA") if im.mode != "RGBA" else im
    ImageOps.exif_transpose(im, in_place=True)
    return im


def paste_offset(bbox_size: Size, child_size: Size) -> Tuple[int, int]:
    return (bbox_size.w - child_size.w) // 2, (bbox_size.h - child_size.h) // 2


def fit_image(im: Image.Image, size: Size, mode: FitMode = FitMode.CONTAIN) -> Image.Image:
    canvas = Image.new("RGBA", (size.w, size.h), BG_COLOR)

    if mode == FitMode.NONE:
        placed = im
    elif mode == FitMode.STRETCH:
        placed = im.resize((size.w, size.h), resample=RESAMPLE)
    elif mode == FitMode.COVER:
        placed = ImageOps.fit(im, (size.w, size.h), method=RESAMPLE, centering=(0.5, 0.5))
    elif mode == FitMode.CONTAIN:  
        placed = ImageOps.contain(im, (size.w, size.h), method=RESAMPLE)
    else:
        raise ValueError(f"Unknown fit mode: {mode}")
    
    dw, dh = paste_offset(size, Size(placed.width, placed.height))
    canvas.paste(placed, (dw, dh), placed)
    return canvas


def place_padding(im: Image.Image, pad: Padding) -> Image.Image:
    canvas = Image.new("RGBA", (im.width + pad.left + pad.right, im.height + pad.top + pad.bottom), BG_COLOR)
    canvas.paste(im, (pad.left, pad.top), im)
    return canvas
