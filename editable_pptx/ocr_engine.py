from __future__ import annotations

from PIL import Image

from editable_pptx.model import OcrLine


def ensure_mta() -> None:
    try:
        from winrt.runtime import ApartmentType, init_apartment
    except Exception:
        return
    try:
        init_apartment(ApartmentType.MULTI_THREADED)
    except Exception:
        pass


def query_languages() -> list[tuple[str, str]]:
    from winrt.windows.media.ocr import OcrEngine

    languages = []
    for language in OcrEngine.available_recognizer_languages:
        languages.append((language.language_tag, language.display_name))
    return languages


def _is_cjk(ch: str) -> bool:
    if not ch:
        return False
    code = ord(ch)
    return (
        0x3400 <= code <= 0x9FFF
        or 0x3000 <= code <= 0x303F
        or 0xFF00 <= code <= 0xFFEF
    )


def join_ocr_words(words: list[str]) -> str:
    pieces: list[str] = []
    previous = ""
    for word in words:
        if not word:
            continue
        if previous:
            left, right = previous[-1], word[0]
            if left.isalnum() and right.isalnum() and not _is_cjk(left) and not _is_cjk(right):
                pieces.append(" ")
        pieces.append(word)
        previous = word
    return "".join(pieces).strip()


def _clean_text(text: str) -> str:
    return "".join(ch for ch in text if ch >= " " or ch == "\n").strip()


def lines_from_result(result: dict) -> list[OcrLine]:
    lines: list[OcrLine] = []
    for line in result.get("lines") or []:
        words = line.get("words") or []
        word_text = [str(word.get("text") or "") for word in words]
        text = join_ocr_words(word_text) if any(word_text) else str(line.get("text") or "")
        text = _clean_text(text)
        if not text:
            continue
        rects = []
        for word in words:
            rect = word.get("bounding_rect") or {}
            if "x" not in rect or "width" not in rect:
                continue
            rects.append(
                (
                    float(rect["x"]),
                    float(rect["y"]),
                    float(rect["width"]),
                    float(rect["height"]),
                )
            )
        if not rects:
            continue
        x0 = min(item[0] for item in rects)
        y0 = min(item[1] for item in rects)
        x1 = max(item[0] + item[2] for item in rects)
        y1 = max(item[1] + item[3] for item in rects)
        lines.append(OcrLine(text, x0, y0, max(1.0, x1 - x0), max(1.0, y1 - y0)))
    return lines


def _downscale(image: Image.Image, max_edge: int = 2200) -> tuple[Image.Image, float]:
    edge = max(image.size)
    if edge <= max_edge:
        return image, 1.0
    scale = max_edge / edge
    resized = image.resize(
        (max(1, int(image.width * scale)), max(1, int(image.height * scale))),
        Image.Resampling.LANCZOS,
    )
    return resized, scale


def recognize_lines(image: Image.Image, lang: str) -> list[OcrLine]:
    if not lang:
        raise RuntimeError("尚未選擇文字辨識語言。")
    import winocr
    from winrt.windows.globalization import Language
    from winrt.windows.media.ocr import OcrEngine

    if not OcrEngine.is_language_supported(Language(lang)):
        raise RuntimeError(
            f"這台電腦沒有安裝「{lang}」的光學字元辨識。"
            "請到 設定 → 時間與語言 → 語言與地區，開啟該語言的語言選項，安裝光學字元辨識。"
        )
    work, scale = _downscale(image)
    try:
        result = winocr.recognize_pil_sync(work, lang)
    except AssertionError as exc:
        raise RuntimeError("無法啟動 Windows 文字辨識。請確認已安裝對應語言的光學字元辨識。") from exc
    lines = lines_from_result(result)
    if scale != 1.0:
        factor = 1.0 / scale
        for line in lines:
            line.x *= factor
            line.y *= factor
            line.w *= factor
            line.h *= factor
    return lines
