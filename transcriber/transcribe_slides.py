"""Utilities for extracting and transcribing slide audio."""

import os
import tempfile
import zipfile
from xml.etree import ElementTree as ET
from collections import defaultdict

from pptx import Presentation
from openai import OpenAI


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
            # Count audio tracks for this slide
            audio_tracks = 0
            for rel in tree.getroot():
                if "audio" in rel.attrib.get("Type", "").lower():
                    audio_tracks += 1

            for track_idx, rel in enumerate(tree.getroot(), start=1):
                rel_type = rel.attrib.get("Type", "").lower()
                target = rel.attrib.get("Target")
                if "audio" not in rel_type or not target:
                    continue
                # Get base directory from the parent XML of the .rels file
                parent_xml = rels_name.replace('_rels/', '').replace('.rels', '')
                base_dir = os.path.dirname(parent_xml)  # e.g., ppt/slides
                # Handle relative paths properly
                if target.startswith('../'):
                    # Go up one level from base_dir and append the rest of the path
                    media_path = os.path.join(os.path.dirname(base_dir), target[3:]).replace('\\', '/')
                else:
                    media_path = os.path.join(base_dir, target).replace('\\', '/')
                if media_path not in z.namelist():
                    continue
                ext = os.path.splitext(media_path)[1]
                out_path = os.path.join(temp_dir, f"slide{idx}_track{track_idx}{ext}")
                with z.open(media_path) as media_file, open(out_path, "wb") as out:
                    out.write(media_file.read())
                yield idx, track_idx, out_path, audio_tracks


def transcribe_slides(pptx_path: str, output_dir: str, model: str = "whisper-1", language: str = "en",
                     task: str = "transcribe", prefix: str = "slide", api_key: str = None):
    """Transcribe audio from PowerPoint slides."""
    os.makedirs(output_dir, exist_ok=True)
    client = OpenAI(api_key=api_key or os.getenv("OPENAI_API_KEY"))
    if not client.api_key:
        print("OpenAI API key required. Use --api-key or set OPENAI_API_KEY.")
        return

    with tempfile.TemporaryDirectory() as tmpdir:
        # Group audio files by slide
        slide_files = defaultdict(list)
        for idx, track_idx, audio_path, total_tracks in extract_slide_audio(pptx_path, tmpdir):
            slide_files[idx].append((track_idx, audio_path, total_tracks))

        if not slide_files:
            print("No audio found in PPTX")
            return

        # Create a single output file
        out_file = os.path.join(output_dir, f"{prefix}_transcription.txt")
        with open(out_file, "w", encoding="utf-8") as f:
            # Process each slide's audio files
            for slide_idx, tracks in sorted(slide_files.items()):
                total_tracks = tracks[0][2]  # Get total tracks from first entry
                f.write(f"\n{'='*50}\n")
                f.write(f"Slide {slide_idx}\n")
                f.write(f"{'='*50}\n\n")

                for track_idx, audio_path, _ in sorted(tracks):
                    with open(audio_path, "rb") as audio_file:
                        if task == "translate":
                            result = client.audio.translations.create(
                                model=model,
                                file=audio_file,
                                language=language,
                            )
                        else:
                            result = client.audio.transcriptions.create(
                                model=model,
                                file=audio_file,
                                language=language,
                            )
                    text = result.text.strip()
                    if total_tracks > 1:
                        f.write(f"Track {track_idx} of {total_tracks}:\n")
                    f.write(text)
                    f.write("\n\n")
                print(f"Transcribed slide {slide_idx}")
            print(f"\nAll transcriptions saved to: {out_file}")


def main(argv=None):
    """Run the application using command-line arguments."""

    from transcriber.cli import build_parser
    parser = build_parser()
    if argv is None:
        argv = os.sys.argv[1:]

    if "--gui" in argv:
        argv.remove("--gui")

        from gooey import Gooey

        @Gooey(program_name="Slides Transcriber")
        def _gui_main():
            args = parser.parse_args(argv)
            transcribe_slides(
                args.pptx,
                args.output,
                model=args.model,
                language=args.language,
                task=args.task,
                prefix=args.prefix,
                api_key=args.api_key
            )

        _gui_main()
    else:
        args = parser.parse_args(argv)
        transcribe_slides(
            args.pptx,
            args.output,
            model=args.model,
            language=args.language,
            task=args.task,
            prefix=args.prefix,
            api_key=args.api_key
        )


if __name__ == "__main__":
    main()
