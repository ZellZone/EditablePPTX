from __future__ import annotations

from PIL import Image

from editable_pptx.imaging import expand_box, fit_block_font, prepare_image, sample_colors
from editable_pptx.model import ConvertOptions, OcrLine, SlideSize, TextBlock
from editable_pptx.ocr_engine import recognize_lines


def block_from_line(
    line: OcrLine,
    image: Image.Image,
    placement: tuple[float, float, float, float],
    slide_size: SlideSize,
) -> TextBlock | None:
    text = line.text.strip()
    if not text:
        return None
    img_w, img_h = image.size
    x = max(0.0, min(float(line.x), img_w - 1))
    y = max(0.0, min(float(line.y), img_h - 1))
    w = max(1.0, min(float(line.w), img_w - x))
    h = max(1.0, min(float(line.h), img_h - y))
    if w < 8 or h < 8:
        return None
    fg, bg = sample_colors(image, (x, y, w, h))
    nx, ny, nw, nh = placement
    sx = nx + (x / img_w) * nw
    sy = ny + (y / img_h) * nh
    sw = max(0.004, (w / img_w) * nw)
    sh = max(0.004, (h / img_h) * nh)
    sy = max(0.0, sy - sh * 0.08)
    sh = min(1.0 - sy, sh * 1.2)
    sw = min(1.0 - sx, sw * 1.12)
    if sh < 0.008 or sw < 0.008:
        return None
    block = TextBlock(
        text=text,
        x=sx,
        y=sy,
        w=sw,
        h=sh,
        color=fg,
        bg=bg,
        cover_box=expand_box(x, y, w, h, img_w, img_h),
    )
    fit_block_font(block, slide_size)
    return block


def analyze_image(image: Image.Image, options: ConvertOptions, slide_size: SlideSize) -> list[TextBlock]:
    if options.mode != "text":
        return []
    prepared, placement = prepare_image(image, slide_size, options.fit)
    blocks: list[TextBlock] = []
    for line in recognize_lines(prepared, options.lang):
        block = block_from_line(line, prepared, placement, slide_size)
        if block is not None:
            blocks.append(block)
    blocks.sort(key=lambda item: (round(item.y, 3), item.x))
    return blocks
