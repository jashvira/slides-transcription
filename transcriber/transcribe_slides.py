import os
import tempfile
import zipfile
from argparse import Namespace

from gooey import Gooey
from pptx import Presentation
import whisper

from .cli import build_parser


def extract_slide_audio(pptx_path: str, temp_dir: str):
    """Extract audio files from a PPTX and yield (index, path) tuples."""
    prs = Presentation(pptx_path)
    z = zipfile.ZipFile(pptx_path)
    for idx, slide in enumerate(prs.slides, start=1):
        rels_name = f"ppt/slides/_rels/slide{idx}.xml.rels"
        if rels_name not in z.namelist():
            continue
        with z.open(rels_name) as rels_file:
            content = rels_file.read().decode("utf-8")
        # Very naive search for media targets
        for line in content.splitlines():
            if "media" in line and "Target" in line:
                target = line.split("Target=")[-1].split('"')[1]
                if not target.startswith(".."):
                    media_path = f"ppt/slides/{target}"
                else:
                    media_path = target.replace("../", "ppt/")
                if media_path not in z.namelist():
                    continue
                ext = os.path.splitext(media_path)[1]
                out_path = os.path.join(temp_dir, f"slide{idx}{ext}")
                with z.open(media_path) as media_file, open(out_path, "wb") as out:
                    out.write(media_file.read())
                yield idx, out_path


def run(args: Namespace):
    os.makedirs(args.output, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmpdir:
        audio_files = list(extract_slide_audio(args.pptx, tmpdir))
        if not audio_files:
            print("No audio found in PPTX")
            return
        model = whisper.load_model(args.model)
        for idx, audio_path in audio_files:
            result = model.transcribe(audio_path)
            text = result.get("text", "").strip()
            out_file = os.path.join(args.output, f"slide{idx}.txt")
            with open(out_file, "w", encoding="utf-8") as f:
                f.write(text)
            print(f"Transcribed slide {idx} -> {out_file}")


@Gooey(program_name="Slides Transcriber")
def main():
    parser = build_parser()
    args = parser.parse_args()
    run(args)


if __name__ == "__main__":
    main()
