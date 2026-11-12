#!/usr/bin/env python3
"""Apply approved concise notes to the 24-slide defense deck; preserve all visuals and reserves."""
import json
import os
from pathlib import Path
from tempfile import NamedTemporaryFile
from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]
PPTX = ROOT / "Molinario_Thesis_Defense_2026.pptx"
NOTES = ROOT / "defense_speaker_notes_12min.json"

def text(slide):
    return tuple(s.text for s in slide.shapes if s.has_text_frame)

def note(slide):
    return slide.notes_slide.notes_text_frame.text.strip()

def main():
    entries = json.loads(NOTES.read_text(encoding="utf-8"))["slides"]
    deck = Presentation(PPTX)
    if len(deck.slides) != 24:
        raise RuntimeError("Expected 24 slides; aborting")
    for i, slide in enumerate(deck.slides, 1):
        print(f"CURRENT SLIDE {i:02d}: {' '.join(text(slide))[:110]!r}; show={slide._element.get('show')!r}", flush=True)
    visible = [(text(s), s._element.get("show")) for s in deck.slides]
    protected = {n: note(deck.slides[n-1]) for n in [11, 13, *range(16, 25)]}
    for number, item in entries.items():
        n = int(number)
        slide = deck.slides[n-1]
        if item["title"].lower() not in " ".join(text(slide)).lower():
            raise RuntimeError(f"Unexpected slide {n} title")
        slide.notes_slide.notes_text_frame.text = item["notes"].strip()
    with NamedTemporaryFile(dir=PPTX.parent, prefix=".notes_", suffix=".pptx", delete=False) as f:
        tmp = Path(f.name)
    try:
        deck.save(tmp)
        check = Presentation(tmp)
        if len(check.slides) != 24:
            raise RuntimeError("Slide count changed")
        for i, slide in enumerate(check.slides):
            if (text(slide), slide._element.get("show")) != visible[i]:
                raise RuntimeError(f"Slide {i+1} visual or visibility changed")
        for n, original in protected.items():
            if note(check.slides[n-1]) != original:
                raise RuntimeError(f"Protected notes changed on slide {n}")
        for number, item in entries.items():
            if note(check.slides[int(number)-1]) != item["notes"].strip():
                raise RuntimeError(f"New notes failed verification on slide {number}")
        os.replace(tmp, PPTX)
    finally:
        if tmp.exists():
            tmp.unlink()
    print(f"Updated and verified {len(entries)} slide notes.")

if __name__ == "__main__":
    main()
