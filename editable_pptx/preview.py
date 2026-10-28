from __future__ import annotations

from PIL import Image, ImageDraw

from editable_pptx.imaging import compose, get_font, prepare_image
from editable_pptx.model import SlideSize, TextBlock


def render_preview(
    image: Image.Image,
    blocks: list[TextBlock] | None,
    slide_size: SlideSize,
    fit: str,
    cover_text: bool,
    show_boxes: bool,
    width: int = 1600,
) -> Image.Image:
    width = max(320, int(width))
    height = max(1, int(round(width * slide_size.inches_h / slide_size.inches_w)))
    prepared, placement = prepare_image(image, slide_size, fit)
    drawing = blocks or []
    composed = compose(prepared, drawing, cover_text and blocks is not None)
    canvas = Image.new("RGB", (width, height), (255, 255, 255))
    nx, ny, nw, nh = placement
    dest_w = max(1, int(round(nw * width)))
    dest_h = max(1, int(round(nh * height)))
    resized = composed.resize((dest_w, dest_h), Image.Resampling.LANCZOS)
    dest_x = int(round(nx * width))
    dest_y = int(round(ny * height))
    canvas.paste(resized, (dest_x, dest_y))
    if not blocks:
        return canvas
    draw = ImageDraw.Draw(canvas)
    outline = max(1, width // 700)
    for block in blocks:
        if not block.text.strip():
            continue
        bx = block.x * width
        by = block.y * height
        bw = max(1.0, block.w * width)
        bh = max(1.0, block.h * height)
        font_px = max(1, int(round(block.font_pt * height / (slide_size.inches_h * 72))))
        font = get_font(font_px)
        bbox = font.getbbox(block.text)
        tx = bx - bbox[0]
        text_h = bbox[3] - bbox[1]
        ty = by + (bh - text_h) / 2 - bbox[1]
        draw.text((tx, ty), block.text, font=font, fill=block.color)
        if show_boxes:
            draw.rectangle([bx, by, bx + bw, by + bh], outline=(196, 92, 38), width=outline)
    return canvas
