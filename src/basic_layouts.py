import math
from typing import List, Optional, Sequence, Tuple, Union
from PIL import Image, ImageDraw, ImageFont
from abc import ABC, abstractmethod

from src.images import FitMode, Size, fit_image, formalize_padding, load_image, place_padding, paste_offset
from src.settings import BG_COLOR


def _split_sizes(total: int, parts: int, weights: Sequence[float]) -> List[int]:
    """Integer split of 'total' into 'parts' sizes using weights; sums exactly to total."""
    if parts <= 0:
        return []
    wsum = sum(weights) if weights else 0.0
    if wsum <= 0:
        weights = [1.0] * parts
        wsum = float(parts)

    raw = [total * (w / wsum) for w in weights]
    ints = [int(x) for x in raw]

    # distribute remainder
    rem = total - sum(ints)
    frac_order = sorted(range(parts), key=lambda i: (raw[i] - ints[i]), reverse=True)
    for i in range(rem):
        ints[frac_order[i % parts]] += 1
    return ints


class Layout(ABC):
    """The Base Meta-Class"""
    @abstractmethod
    def natural_size(self) -> Size:
        ...
    
    @abstractmethod
    def render(self, size: Optional[Size] = None) -> Image.Image:
        ...

    @abstractmethod
    def place(self, size: Optional[Size] = None) -> List[Tuple[int, int, Image.Image]]:
        ...


class BlankLayout(Layout):
    def __init__(self, size: Optional[Size] = None, *args, **kwargs):
        self.size = size or Size(0, 0)
    
    def natural_size(self) -> Size:
        return self.size
    
    def render(self, size: Optional[Size] = None) -> Image.Image:
        size = size or self.natural_size()
        return Image.new("RGBA", (size.w, size.h), BG_COLOR)
    
    def place(self, size: Optional[Size] = None) -> List[Tuple[int, int, Image.Image]]:
        return []


class SingleImageLayout(Layout):
    def __init__(
            self,
            image: Union[str, Image.Image],
            *,

            padding: Union[int, Tuple[int, int], Tuple[int, int, int, int]] = 0,

            fit: FitMode = FitMode.CONTAIN,
    ):
        self.image = load_image(image) if isinstance(image, str) else image
        self.padding = formalize_padding(padding)
        self.fit = fit  

    def natural_size(self) -> Size:
        return Size(self.image.width, self.image.height).expand(self.padding)
    
    def render(self, size: Optional[Size] = None) -> Image.Image:
        size = size or self.natural_size()
        bbox = size.inset(self.padding)
        return place_padding(fit_image(self.image, bbox, self.fit), self.padding)
    
    def place(self, size: Optional[Size] = None) -> List[Tuple[int, int, Image.Image]]:
        size = size or self.natural_size()
        bbox = size.inset(self.padding)
        img = fit_image(self.image, bbox, self.fit)
        return [(self.padding.left, self.padding.top, img),]


class SingleTextLayout(Layout):
    def __init__(
            self,
            text: str,
            *,

            padding: Union[int, Tuple[int, int], Tuple[int, int, int, int]] = 0,
            font_size: int = 32,
            color: str = "black",
            align: str = "left",  # left, center, right

            fit: FitMode = FitMode.CONTAIN,
    ):
        self.text = text
        self.padding = formalize_padding(padding)
        self.font_size = font_size
        self.color = color
        self.align = align
        self.fit = fit
    
    def natural_size(self) -> Size:
        dummy = Image.new("RGBA", (1, 1), BG_COLOR)
        d = ImageDraw.Draw(dummy)
        l, t, r, b = d.multiline_textbbox(
            (0, 0), self.text, align=self.align
        )
        w, h = math.ceil((r - l) + self.padding.left + self.padding.right), math.ceil((b - t) + self.padding.top + self.padding.bottom)
        return Size(max(1, w), max(1, h))
    
    def render(self, size: Optional[Size] = None) -> Image.Image:
        size = size or self.natural_size()
        print(size.w, size.h)
        canvas = Image.new("RGBA", (size.w, size.h), BG_COLOR)
        d = ImageDraw.Draw(canvas)
        d.multiline_text(
            (self.padding.left, self.padding.top), self.text, align=self.align, fill=self.color
        )
        return canvas
    
    def place(self, size: Optional[Size] = None) -> List[Tuple[int, int, Image.Image]]:
        size = size or self.natural_size()
        return [(0, 0, self.render(size))]


class HorizontalLayout(Layout):
    def __init__(
            self,
            children: Sequence[Layout],
            *,

            gaps: Union[int, Sequence[int]] = 0,
            padding: Union[int, Tuple[int, int], Tuple[int, int, int, int]] = 0,
            weights: Optional[Sequence[float]] = None,
    ):
        self.children = children
        self.gaps = gaps if isinstance(gaps, Sequence) else [gaps] * (len(children) - 1)
        self.padding = formalize_padding(padding)
        self.weights = weights or [c.natural_size().w for c in self.children]

    def natural_size(self) -> Size:
        if not self.children:
            return Size(self.padding.left + self.padding.right, self.padding.top + self.padding.bottom)
        sizes = [c.natural_size() for c in self.children]
        w = sum(s.w for s in sizes) + sum(self.gaps) + self.padding.left + self.padding.right
        h = max(s.h for s in sizes) + self.padding.top + self.padding.bottom
        return Size(w, h)
    
    def render(self, size: Optional[Size] = None) -> Image.Image:
        size = size or self.natural_size()
        
        bbox = size.inset(self.padding)
        available_w = max(1, bbox.w - sum(self.gaps))
        available_h = max(1, bbox.h)
        chidren_widths = _split_sizes(available_w, len(self.children), self.weights)

        canvas = Image.new("RGBA", (size.w, size.h), BG_COLOR)
        x = self.padding.left
        last_index = len(self.children) - 1
        for i, (child, child_width) in enumerate(zip(self.children, chidren_widths)):
            child_img = child.render(Size(child_width, available_h))
            dw, dh = paste_offset(Size(child_width, available_h), Size(child_img.width, child_img.height))
            canvas.paste(child_img, (x + dw, self.padding.top + dh), child_img)
            x += child_width
            if i < last_index:
                x += self.gaps[i]
        return canvas
    
    def place(self, size: Optional[Size] = None) -> List[Tuple[int, int, Image.Image]]:
        size = size or self.natural_size()

        bbox = size.inset(self.padding)
        available_w = max(1, bbox.w - sum(self.gaps))
        available_h = max(1, bbox.h)
        chidren_widths = _split_sizes(available_w, len(self.children), self.weights)

        placements: List[Tuple[int, int, Image.Image]] = []
        x = self.padding.left
        last_index = len(self.children) - 1
        for i, (child, child_width) in enumerate(zip(self.children, chidren_widths)):
            for cx, cy, cimg in child.place(Size(child_width, available_h)):
                placements.append((x + cx, self.padding.top + cy, cimg))
            x += child_width
            if i < last_index:
                x += self.gaps[i]
        return placements
    

class VerticalLayout(Layout):
    def __init__(
            self,
            children: Sequence[Layout],
            *,

            gaps: Union[int, Sequence[int]] = 0,
            padding: Union[int, Tuple[int, int], Tuple[int, int, int, int]] = 0,
            weights: Optional[Sequence[float]] = None,
    ):
        self.children = children
        self.gaps = gaps if isinstance(gaps, Sequence) else [gaps] * (len(children) - 1)
        self.padding = formalize_padding(padding)
        self.weights = weights or [c.natural_size().h for c in self.children]

    def natural_size(self) -> Size:
        if not self.children:
            return Size(self.padding.left + self.padding.right, self.padding.top + self.padding.bottom)
        sizes = [c.natural_size() for c in self.children]
        w = max(s.w for s in sizes) + self.padding.left + self.padding.right
        h = sum(s.h for s in sizes) + sum(self.gaps) + self.padding.top + self.padding.bottom
        return Size(w, h)
    
    def render(self, size: Optional[Size] = None) -> Image.Image:
        size = size or self.natural_size()
        
        bbox = size.inset(self.padding)
        available_w = max(1, bbox.w)
        available_h = max(1, bbox.h - sum(self.gaps))
        chidren_heights = _split_sizes(available_h, len(self.children), self.weights)

        canvas = Image.new("RGBA", (size.w, size.h), BG_COLOR)
        y = self.padding.top
        last_index = len(self.children) - 1
        for i, (child, child_height) in enumerate(zip(self.children, chidren_heights)):
            child_img = child.render(Size(available_w, child_height))
            dw, dh = paste_offset(Size(available_w, child_height), Size(child_img.width, child_img.height))
            canvas.paste(child_img, (self.padding.left + dw, y + dh), child_img)
            y += child_height
            if i < last_index:
                y += self.gaps[i]
        return canvas
    
    def place(self, size: Optional[Size] = None) -> List[Tuple[int, int, Image.Image]]:
        size = size or self.natural_size()

        bbox = size.inset(self.padding)
        available_w = max(1, bbox.w)
        available_h = max(1, bbox.h - sum(self.gaps))
        chidren_heights = _split_sizes(available_h, len(self.children), self.weights)

        placements: List[Tuple[int, int, Image.Image]] = []
        y = self.padding.top
        last_index = len(self.children) - 1
        for i, (child, child_height) in enumerate(zip(self.children, chidren_heights)):
            for cx, cy, cimg in child.place(Size(available_w, child_height)):
                placements.append((self.padding.left + cx, y + cy, cimg))
            y += child_height
            if i < last_index:
                y += self.gaps[i]
        return placements


class AdaptiveGridLayout(Layout):
    """automatically arranges children in a grid of variable size, with equal weights"""
    def __init__(
            self,
            children: Sequence[Layout],
            *,

            row_gap: int = 0,
            col_gap: int = 0,
            padding: Union[int, Tuple[int, int], Tuple[int, int, int, int]] = 0,
    ):
        self.children = children
        self.row_gap = row_gap
        self.col_gap = col_gap
        self.padding = formalize_padding(padding)

    def _adaptive_grid(self, N: int, cell_aspect_ratio: float, render_aspect_ratio: float) -> Tuple[int, int]:
        if N <= 0:
            return 0, 0
        n_cols = math.ceil(math.sqrt(N * cell_aspect_ratio / render_aspect_ratio))
        n_rows = math.ceil(N / n_cols)
        return n_cols, n_rows
    
    def natural_size(self) -> Size:
        if not self.children:
            return Size(self.padding.left + self.padding.right, self.padding.top + self.padding.bottom)
        n_cols, n_rows = self._adaptive_grid(len(self.children), 1.0, 1.0)
        sizes = [c.natural_size() for c in self.children]
        w = max(s.w for s in sizes) * n_cols + self.col_gap * (n_cols - 1) + self.padding.left + self.padding.right
        h = max(s.h for s in sizes) * n_rows + self.row_gap * (n_rows - 1) + self.padding.top + self.padding.bottom
        return Size(w, h)
    
    def render(self, size: Optional[Size] = None) -> Image.Image:
        size = size or self.natural_size()
        if not self.children:
            return Image.new("RGBA", (size.w, size.h), BG_COLOR)
        
        bbox = size.inset(self.padding)
        sizes = [c.natural_size() for c in self.children]
        max_child_aspect_ratio = max(s.w for s in sizes) / max(s.h for s in sizes)
        n_cols, n_rows = self._adaptive_grid(len(self.children), 
            cell_aspect_ratio=max_child_aspect_ratio, render_aspect_ratio=bbox.w / bbox.h)
        cell_w = (bbox.w - self.col_gap * (n_cols - 1)) // n_cols
        cell_h = (bbox.h - self.row_gap * (n_rows - 1)) // n_rows
        
        canvas = Image.new("RGBA", (size.w, size.h), BG_COLOR)
        y = self.padding.top
        for i in range(n_rows):
            x = self.padding.left
            for j in range(n_cols):
                idx = i * n_cols + j
                if idx >= len(self.children):
                    break
                child = self.children[idx]
                child_img = child.render(Size(cell_w, cell_h))
                dw, dh = paste_offset(Size(cell_w, cell_h), Size(child_img.width, child_img.height))
                canvas.paste(child_img, (x + dw, y + dh), child_img)
                x += cell_w + self.col_gap
            y += cell_h + self.row_gap
        return canvas
    
    def place(self, size: Optional[Size] = None) -> List[Tuple[int, int, Image.Image]]:
        size = size or self.natural_size()
        if not self.children:
            return []

        bbox = size.inset(self.padding)
        sizes = [c.natural_size() for c in self.children]
        max_child_aspect_ratio = max(s.w for s in sizes) / max(s.h for s in sizes)
        n_cols, n_rows = self._adaptive_grid(
            len(self.children), cell_aspect_ratio=max_child_aspect_ratio, render_aspect_ratio=bbox.w / bbox.h
        )
        cell_w = (bbox.w - self.col_gap * (n_cols - 1)) // n_cols
        cell_h = (bbox.h - self.row_gap * (n_rows - 1)) // n_rows

        placements: List[Tuple[int, int, Image.Image]] = []
        y = self.padding.top
        for i in range(n_rows):
            x = self.padding.left
            for j in range(n_cols):
                idx = i * n_cols + j
                if idx >= len(self.children):
                    break
                child = self.children[idx]
                for cx, cy, cimg in child.place(Size(cell_w, cell_h)):
                    placements.append((x + cx, y + cy, cimg))
                x += cell_w + self.col_gap
            y += cell_h + self.row_gap
        return placements


