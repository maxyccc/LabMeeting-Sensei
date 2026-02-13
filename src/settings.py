from PIL import Image

from typing import Union, Tuple


SLIDE_WIDTH = 11
SLIDE_HEIGHT = 8.5

BG_COLOR: Union[str, Tuple[int, int, int], Tuple[int, int, int, int]] = (255, 255, 255, 0)

RESAMPLE = Image.Resampling.LANCZOS