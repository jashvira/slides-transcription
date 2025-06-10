import os
import tempfile
import zipfile
from argparse import Namespace
from xml.etree import ElementTree as ET

from gooey import Gooey
from pptx import Presentation
import whisper

from .cli import build_parser


def extract_slide_audio(pptx_path: str, temp_dir: str):
    """Extract audio files from a PPTX and yield (slide, track, path) tuples."""
    prs = Presentation(pptx_path)
    with zipfile.ZipFile(pptx_path) as z:
        for idx, _ in enumerate(prs.slides, start=1):
            rels_name = f"ppt/slides/_rels/slide{idx}.xml.rels"
            if rels_name not in z.namelist():
                continue
            with z.open(rels_name) as rels_file:
                tree = ET.parse(rels_file)
            for track_idx, rel in enumerate(tree.getroot(), start=1):
                rel_type = rel.attrib.get("Type", "").lower()
                if "audio" not in rel_type:
                    continue
                target = rel.attrib.get("Target")
                if not target:
                    continue
                rel_dir = os.path.dirname(rels_name)
                media_path = os.path.normpath(os.path.join(rel_dir, target))
                if media_path not in z.namelist():
                    continue
                ext = os.path.splitext(media_path)[1]
                out_path = os.path.join(temp_dir, f"slide{idx}_{track_idx}{ext}")
                with z.open(media_path) as media_file, open(out_path, "wb") as out:
                    out.write(media_file.read())
                yield idx, track_idx, out_path


def run(args: Namespace):
    os.makedirs(args.output, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmpdir:
        audio_files = list(extract_slide_audio(args.pptx, tmpdir))
        if not audio_files:
            print("No audio found in PPTX")
            return
        model = whisper.load_model(args.model)
        for idx, track_idx, audio_path in audio_files:
            result = model.transcribe(
                audio_path,
                language=args.language,
                task=args.task,
            )
            text = result.get("text", "").strip()
            out_name = f"{args.prefix}{idx}_{track_idx}.txt"
            out_file = os.path.join(args.output, out_name)
            with open(out_file, "w", encoding="utf-8") as f:
                f.write(text)
            print(f"Transcribed slide {idx} track {track_idx} -> {out_file}")


def main(argv=None):
    parser = build_parser()
    if argv is None:
        argv = os.sys.argv[1:]

    if "--gui" in argv:
        argv.remove("--gui")

        @Gooey(program_name="Slides Transcriber")
        def _gui_main():
            args = parser.parse_args(argv)
            run(args)

        _gui_main()
    else:
        args = parser.parse_args(argv)
        run(args)


if __name__ == "__main__":
    main()
