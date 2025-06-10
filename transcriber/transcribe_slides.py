"""Utilities for extracting and transcribing slide audio."""

import os
import tempfile
import zipfile
from xml.etree import ElementTree as ET
from collections import defaultdict
import sys
import shutil
import mimetypes
import time
import subprocess
import zlib
import re
from tqdm import tqdm

from pptx import Presentation
from openai import OpenAI, OpenAIError

# Supported audio formats by OpenAI
SUPPORTED_AUDIO_FORMATS = {
    'audio/wav', 'audio/mpeg', 'audio/mp4', 'audio/mpeg', 'audio/mpga',
    'audio/webm', 'audio/x-m4a'
}

# PowerPoint audio file extensions and their corresponding MIME types
AUDIO_EXTENSIONS = {
    '.wav': 'audio/wav',
    '.mp3': 'audio/mpeg',
    '.m4a': 'audio/x-m4a',
    '.mp4': 'audio/mp4',
    '.mpeg': 'audio/mpeg',
    '.mpga': 'audio/mpga',
    '.webm': 'audio/webm',
    '.bin': 'audio/wav'  # PowerPoint .bin files are typically WAV files
}

def get_file_size_mb(file_path):
    """Get file size in megabytes."""
    return os.path.getsize(file_path) / (1024 * 1024)

def convert_bin_to_wav(bin_path):
    """Convert PowerPoint .bin audio file to WAV format using ffmpeg."""
    try:
        wav_path = os.path.splitext(bin_path)[0] + '.wav'
        # Use ffmpeg to convert the .bin file to WAV
        # PowerPoint .bin files are typically WAV files with a different extension
        subprocess.run(['ffmpeg', '-i', bin_path, '-acodec', 'pcm_s16le', '-ar', '44100', wav_path],
                      check=True, capture_output=True)
        return wav_path
    except subprocess.CalledProcessError as e:
        raise ValueError(f"Failed to convert audio file: {e.stderr.decode()}")
    except Exception as e:
        raise ValueError(f"Failed to convert audio file: {str(e)}")

def validate_audio_file(file_path):
    """Validate if the audio file is in a supported format and not corrupted."""
    try:
        # Check file extension first
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in AUDIO_EXTENSIONS:
            return False, f"Unsupported audio extension: {ext}"

        # For .bin files, we'll convert them to WAV
        if ext == '.bin':
            try:
                wav_path = convert_bin_to_wav(file_path)
                return True, wav_path
            except Exception as e:
                return False, f"Failed to convert .bin file: {str(e)}"

        # Try MIME type detection for other formats
        mime_type, _ = mimetypes.guess_type(file_path)
        if mime_type and mime_type not in SUPPORTED_AUDIO_FORMATS:
            return False, f"Unsupported audio format: {mime_type}"

        # Basic file integrity check
        if os.path.getsize(file_path) == 0:
            return False, "Audio file is empty"

        return True, None
    except Exception as e:
        return False, f"Failed to validate audio file: {str(e)}"

def validate_api_key(api_key):
    """Validate the OpenAI API key."""
    try:
        client = OpenAI(api_key=api_key)
        # Make a minimal API call to validate the key
        client.models.list()
        return True, None
    except OpenAIError as e:
        return False, f"Invalid API key: {str(e)}"
    except Exception as e:
        return False, f"API validation error: {str(e)}"

def check_disk_space(directory, required_space_mb):
    """Check if there's enough disk space."""
    try:
        free_space = shutil.disk_usage(directory).free / (1024 * 1024)  # Convert to MB
        return free_space >= required_space_mb, free_space
    except Exception as e:
        return False, f"Failed to check disk space: {str(e)}"

def check_pptx_integrity(pptx_path: str) -> tuple[bool, list[str]]:
    """Check if PPTX file is valid and not corrupted.
    Returns (is_valid, corrupted_files) tuple."""
    try:
        # First try to open as ZIP
        with zipfile.ZipFile(pptx_path) as z:
            # Get all files
            all_files = z.namelist()

            # Check essential PPTX files
            required_files = [
                '[Content_Types].xml',
                '_rels/.rels',
                'ppt/presentation.xml',
                'ppt/_rels/presentation.xml.rels'
            ]

            missing_files = []
            for file in required_files:
                if file not in all_files:
                    missing_files.append(file)

            if missing_files:
                return False, missing_files

            # Check file integrity
            corrupted_files = []
            for file in all_files:
                try:
                    z.read(file)
                except zlib.error:
                    corrupted_files.append(file)
                except Exception:
                    corrupted_files.append(file)

            # Check media files specifically
            media_files = [f for f in all_files if f.startswith('ppt/media/')]
            audio_files = [f for f in media_files if any(f.lower().endswith(ext) for ext in AUDIO_EXTENSIONS.keys())]

            if not audio_files:
                return False, corrupted_files

            return True, corrupted_files

    except zipfile.BadZipFile:
        return False, []
    except Exception:
        return False, []

def extract_slide_audio(pptx_path: str, output_dir: str):
    """Extract audio from PowerPoint slides."""
    try:
        # First check file integrity
        is_valid, corrupted_files = check_pptx_integrity(pptx_path)
        if not is_valid:
            if not corrupted_files:
                raise ValueError("PPTX file appears to be corrupted. Please try repairing the file in PowerPoint first.")
            print("\nNote: Some files in the presentation are corrupted. The script will attempt to process the remaining files.")

        # Get total number of slides by counting slide XML files
        with zipfile.ZipFile(pptx_path) as z:
            slide_files = [f for f in z.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')]
            total_slides = len(slide_files)

            print(f"\nProcessing {total_slides} slides...")

            for idx in range(1, total_slides + 1):
                rels_name = f"ppt/slides/_rels/slide{idx}.xml.rels"
                if rels_name not in z.namelist():
                    print(f"Slide {idx}: No relationships file found, skipping.", flush=True)
                    continue

                try:
                    with z.open(rels_name) as rels_file:
                        tree = ET.parse(rels_file)
                        audio_tracks = [rel for rel in tree.getroot() if "audio" in rel.attrib.get("Type", "").lower()]

                        if not audio_tracks:
                            print(f"Slide {idx}: No audio tracks found, skipping.", flush=True)
                            continue

                        for track_idx, rel in enumerate(audio_tracks, 1):
                            target = rel.attrib.get("Target", "")
                            if not target:
                                print(f"Slide {idx} Track {track_idx}: No target found, skipping.", flush=True)
                                continue

                            media_path = target.replace("../", "ppt/")
                            if media_path in corrupted_files:
                                print(f"Slide {idx} Track {track_idx}: Corrupted audio file {media_path}, skipping.", flush=True)
                                continue

                            print(f"Slide {idx} Track {track_idx}: Extracting {media_path}...", flush=True)
                            try:
                                with z.open(media_path) as audio_file:
                                    audio_data = audio_file.read()
                                    ext = os.path.splitext(target)[1].lower()
                                    if ext == ".bin":
                                        ext = ".wav"
                                    output_path = os.path.join(output_dir, f"slide{idx}_track{track_idx}{ext}")

                                    with open(output_path, "wb") as f:
                                        f.write(audio_data)

                                    print(f"Slide {idx} Track {track_idx}: Extraction successful.", flush=True)
                                    yield idx, track_idx, output_path, len(audio_tracks)
                            except Exception as e:
                                print(f"Slide {idx} Track {track_idx}: Exception during extraction: {e}", flush=True)
                                continue
                except Exception as e:
                    print(f"Slide {idx}: Exception opening relationships or parsing XML: {e}", flush=True)
                    continue

    except Exception as e:
        raise ValueError(f"Failed to open PowerPoint file: {str(e)}")

def inspect_pptx_config(pptx_path: str):
    """Inspect PPTX file configuration and print details about audio files."""
    print(f"\nInspecting PPTX file: {os.path.basename(pptx_path)}")
    print("=" * 50)

    try:
        with zipfile.ZipFile(pptx_path) as z:
            # Get all media files
            media_files = [f for f in z.namelist() if f.startswith('ppt/media/')]
            audio_files = [f for f in media_files if any(f.lower().endswith(ext) for ext in AUDIO_EXTENSIONS.keys())]

            print(f"\nFile Statistics:")
            print(f"  Total media files: {len(media_files)}")
            print(f"  Total audio files: {len(audio_files)}")
            print(f"  PPTX size: {get_file_size_mb(pptx_path):.1f}MB")

            if audio_files:
                print("\nAudio File Types:")
                extensions = {}
                total_audio_size = 0
                for f in audio_files:
                    ext = os.path.splitext(f)[1].lower()
                    extensions[ext] = extensions.get(ext, 0) + 1
                    # Get size of each audio file
                    with z.open(f) as audio_file:
                        total_audio_size += len(audio_file.read()) / (1024 * 1024)  # Convert to MB

                for ext, count in extensions.items():
                    print(f"  {ext}: {count} files")
                print(f"  Total audio size: {total_audio_size:.1f}MB")

            # Check slide relationships
            print("\nSlide Audio Configuration:")
            slides_with_audio = 0
            total_audio_tracks = 0
            max_tracks_per_slide = 0

            for idx in range(1, 1000):  # Check up to 1000 slides
                rels_name = f"ppt/slides/_rels/slide{idx}.xml.rels"
                if rels_name not in z.namelist():
                    break

                with z.open(rels_name) as rels_file:
                    tree = ET.parse(rels_file)
                    audio_tracks = [rel for rel in tree.getroot() if "audio" in rel.attrib.get("Type", "").lower()]
                    if audio_tracks:
                        slides_with_audio += 1
                        total_audio_tracks += len(audio_tracks)
                        max_tracks_per_slide = max(max_tracks_per_slide, len(audio_tracks))
                        print(f"\nSlide {idx}:")
                        print(f"  Audio tracks: {len(audio_tracks)}")
                        for track_idx, rel in enumerate(audio_tracks, 1):
                            target = rel.attrib.get("Target", "")
                            print(f"  Track {track_idx}: {target}")

            print("\nSummary:")
            print(f"  Total slides with audio: {slides_with_audio}")
            print(f"  Total audio tracks: {total_audio_tracks}")
            print(f"  Maximum tracks per slide: {max_tracks_per_slide}")
            print(f"  Average tracks per slide: {total_audio_tracks/slides_with_audio if slides_with_audio else 0:.1f}")

    except Exception as e:
        print(f"Error inspecting PPTX: {str(e)}")

def get_safe_filename(filename: str) -> str:
    """Convert a filename to a safe string that can be used as a prefix."""
    # Remove extension and path
    name = os.path.splitext(os.path.basename(filename))[0]
    # Replace invalid characters with underscores
    safe_name = re.sub(r'[<>:"/\\|?*]', '_', name)
    # Remove leading/trailing spaces and dots
    safe_name = safe_name.strip('. ')
    return safe_name

def estimate_processing_time(audio_files: list, model: str = "whisper-1") -> float:
    """Estimate total processing time in minutes."""
    # Average processing time per file (based on actual measurements)
    if model == "whisper-1":
        return len(audio_files) * 0.083  # Rough estimate: 5 seconds per file
    return len(audio_files) * 0.167  # Conservative estimate for other models (10 seconds)

def estimate_cost(audio_size_mb: float, model: str = "whisper-1") -> float:
    """Estimate cost in USD for transcription."""
    # Current OpenAI pricing (as of 2024)
    if model == "whisper-1":
        return audio_size_mb * 0.006  # $0.006 per minute
    return audio_size_mb * 0.015  # Conservative estimate for other models

def transcribe_slides(
    pptx_path: str,
    output_dir: str,
    model: str = "whisper-1",
    language: str = "en",
    task: str = "transcribe",
    prefix: str = None,
    api_key: str = None,
    progress_callback=None,
    file_num=None,
    total_files=None,
    total_slides_in_file=None
):
    """Transcribe audio from PowerPoint slides."""
    start_time = time.time()

    # Generate prefix from filename if not provided
    if prefix is None:
        prefix = get_safe_filename(pptx_path)
        print(f"Using prefix: {prefix}")

    # Validate API key
    if not api_key:
        raise ValueError("No OpenAI API key provided. Please use --api-key or set the OPENAI_API_KEY environment variable.")

    is_valid, error_msg = validate_api_key(api_key)
    if not is_valid:
        raise ValueError(error_msg)

    # Create and validate output directory
    try:
        os.makedirs(output_dir, exist_ok=True)
        # Test write permissions
        test_file = os.path.join(output_dir, '.write_test')
        with open(test_file, 'w') as f:
            f.write('test')
        os.remove(test_file)
    except Exception as e:
        raise ValueError(f"Output directory error: {str(e)}")

    # Check disk space (estimate 2x the PPTX size for temporary files)
    pptx_size = get_file_size_mb(pptx_path)
    has_space, free_space = check_disk_space(output_dir, pptx_size * 2)
    if not has_space:
        raise ValueError(f"Insufficient disk space. Required: {pptx_size*2:.1f}MB, Available: {free_space:.1f}MB")

    # Check file integrity first
    is_valid, corrupted_files = check_pptx_integrity(pptx_path)
    if not is_valid and not corrupted_files:
        raise ValueError("PPTX file appears to be corrupted. Please try repairing the file in PowerPoint first.")

    # Initialize OpenAI client
    client = OpenAI(api_key=api_key)

    # Create temporary directory for audio files
    with tempfile.TemporaryDirectory() as tmpdir:
        # Extract audio files
        audio_files = []
        try:
            for idx, track_idx, audio_path, total_tracks in extract_slide_audio(pptx_path, tmpdir):
                audio_files.append((idx, track_idx, audio_path, total_tracks))
        except Exception as e:
            raise ValueError(f"Failed to extract audio: {str(e)}")

        if not audio_files:
            print("DEBUG: No audio files found, raising error.", flush=True)
            raise ValueError("No audio files found in the presentation")

        # Estimate processing time and cost
        est_time = estimate_processing_time(audio_files, model)
        total_audio_size = sum(get_file_size_mb(audio_path) for _, _, audio_path, _ in audio_files)
        est_cost = estimate_cost(total_audio_size, model)

        print(f"\nEstimated processing time: {est_time:.1f} minutes")
        print(f"Estimated cost: ${est_cost:.2f}")

        if not sys.stdin.isatty():
            print("GUI mode detected, proceeding without confirmation.", flush=True)
            proceed = 'y'
        else:
            proceed = input("\nDo you want to proceed? (y/n): ").lower()
        if proceed != 'y':
            print("Transcription cancelled.")
            return

        # Create a single output file
        out_file = os.path.join(output_dir, f"{prefix}_transcription.txt")
        with open(out_file, "w", encoding="utf-8") as f:
            # Write summary header
            f.write(f"Transcription Summary\n")
            f.write(f"{'='*50}\n")
            f.write(f"Total audio files: {len(audio_files)}\n")
            f.write(f"Model: {model}\n")
            f.write(f"Language: {language}\n")
            f.write(f"Task: {task}\n")
            f.write(f"Estimated cost: ${est_cost:.2f}\n")
            f.write(f"{'='*50}\n\n")

            # Process each audio file
            total_audio_size = 0
            skipped_files = 0
            failed_slides = set()
            for slide_idx, (idx, track_idx, audio_path, total_tracks) in enumerate(audio_files, 1):
                print(f"Transcribing file {file_num}/{total_files}: {os.path.basename(pptx_path)} (Slide {slide_idx}/{total_slides_in_file})", flush=True)
                try:
                    # Validate audio file
                    is_valid, error_msg = validate_audio_file(audio_path)
                    if not is_valid:
                        skipped_files += 1
                        failed_slides.add(idx)
                        continue

                    # Check file size
                    audio_size = get_file_size_mb(audio_path)
                    total_audio_size += audio_size
                    if audio_size > 25:
                        skipped_files += 1
                        failed_slides.add(idx)
                        continue

                    # Write slide header
                    f.write(f"\n{'='*50}\n")
                    f.write(f"Slide {idx}\n")
                    f.write(f"{'='*50}\n\n")

                    # Transcribe audio
                    print(f"Uploading audio for Slide {slide_idx}...", flush=True)
                    with open(audio_path, "rb") as audio_file:
                        if task == "translate":
                            response = client.audio.translations.create(
                                model=model,
                                file=audio_file,
                                language=language
                            )
                        else:
                            response = client.audio.transcriptions.create(
                                model=model,
                                file=audio_file,
                                language=language
                            )
                    print(f"Transcription successful for Slide {slide_idx}.", flush=True)

                    # Write transcription
                    if total_tracks > 1:
                        f.write(f"Track {track_idx} of {total_tracks}:\n")
                    f.write(response.text.strip())
                    f.write("\n\n")

                except Exception as e:
                    error_msg = f"Failed to transcribe slide {idx}, track {track_idx}: {str(e)}"
                    f.write(f"Error: {error_msg}\n\n")
                    skipped_files += 1
                    failed_slides.add(idx)
                    continue

            # Add final summary
            processing_time = time.time() - start_time
            actual_cost = estimate_cost(total_audio_size, model)
            f.write(f"\n{'='*50}\n")
            f.write("Final Summary\n")
            f.write(f"{'='*50}\n")
            f.write(f"Successfully transcribed: {len(audio_files) - skipped_files} files\n")
            f.write(f"Skipped/Failed: {skipped_files} files\n")
            f.write(f"Failed slides: {', '.join(map(str, sorted(failed_slides))) if failed_slides else 'None'}\n")
            f.write(f"Processing time: {processing_time/60:.1f} minutes\n")
            f.write(f"Estimated cost: ${est_cost:.2f}\n")
            f.write(f"Actual cost: ${actual_cost:.2f}\n")

            print(f"\nTranscription complete! Results saved to: {out_file}")
            print(f"Actual cost: ${actual_cost:.2f}")

def main(argv=None):
    try:
        from transcriber.cli import build_parser
        parser = build_parser()
        if argv is None:
            argv = sys.argv[1:]

        if "--gui" in argv:
            argv.remove("--gui")

            from gooey import Gooey

            @Gooey(program_name="Slides Transcriber")
            def _gui_main():
                args = parser.parse_args(argv)
                # Inspect PPTX first
                inspect_pptx_config(args.pptx)
                # Then proceed with transcription
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
            # Inspect PPTX first
            inspect_pptx_config(args.pptx)
            # Then proceed with transcription
            transcribe_slides(
                args.pptx,
                args.output,
                model=args.model,
                language=args.language,
                task=args.task,
                prefix=args.prefix,
                api_key=args.api_key
            )
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
