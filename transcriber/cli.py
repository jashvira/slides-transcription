"""Command line interface for transcribing slide audio."""

import argparse
import os
from gooey import Gooey, GooeyParser
from transcriber.transcribe_slides import transcribe_slides


def build_parser():
    """Return an argument parser for the CLI."""

    parser = argparse.ArgumentParser(
        description="Transcribe slide audio using the OpenAI Speech to Text API"
    )
    parser.add_argument("pptx", help="Path to PPTX file")
    parser.add_argument("output", help="Output directory for transcripts")
    parser.add_argument(
        "--model",
        help="OpenAI STT model to use (currently only whisper-1 is available)",
        default="whisper-1",
        choices=["whisper-1"],
    )
    parser.add_argument(
        "--language",
        help="Language spoken in the audio (e.g., 'en' for English)",
        default="en"
    )
    parser.add_argument(
        "--task",
        help="Transcription task",
        choices=["transcribe", "translate"],
        default="transcribe"
    )
    parser.add_argument(
        "--prefix",
        help="Prefix for output files",
        default="slide"
    )
    parser.add_argument(
        "--api-key",
        help="OpenAI API key (or set OPENAI_API_KEY environment variable)",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch a GUI instead of the CLI",
    )
    return parser


@Gooey(
    program_name="Slides Transcriber",
    program_description="Extract audio from PowerPoint slides and transcribe them using OpenAI's Whisper model",
    default_size=(800, 600),
    tabbed_groups=True,
    navigation='TABBED'
)
def main():
    """Entry point for the GUI/CLI application."""

    parser = GooeyParser(description="Transcribe audio from PowerPoint slides")

    # Input/Output group
    io_group = parser.add_argument_group("Input/Output", gooey_options={'columns': 1})
    io_group.add_argument(
        "pptx_file",
        help="Path to the PowerPoint file",
        widget="FileChooser",
        gooey_options={'wildcard': "PowerPoint files (*.pptx)|*.pptx"}
    )
    io_group.add_argument(
        "output_dir",
        help="Directory to save transcriptions",
        widget="DirChooser"
    )

    # Transcription Settings group
    trans_group = parser.add_argument_group("Transcription Settings", gooey_options={'columns': 1})
    trans_group.add_argument(
        "--model",
        help="OpenAI STT model to use (currently only whisper-1 is available)",
        default="whisper-1",
        choices=["whisper-1"],
        widget="Dropdown"
    )
    trans_group.add_argument(
        "--language",
        help="Language spoken in the audio (e.g., 'en' for English)",
        default="en"
    )
    trans_group.add_argument(
        "--task",
        help="Transcription task",
        choices=["transcribe", "translate"],
        default="transcribe",
        widget="Dropdown"
    )
    trans_group.add_argument(
        "--prefix",
        help="Prefix for output files",
        default="slide"
    )

    # API Settings group
    api_group = parser.add_argument_group("API Settings", gooey_options={'columns': 1})
    api_group.add_argument(
        "--api-key",
        help="OpenAI API key (or set OPENAI_API_KEY environment variable)",
        default=os.environ.get("OPENAI_API_KEY", ""),
        widget="PasswordField"
    )

    args = parser.parse_args()

    if not args.api_key:
        parser.error("OpenAI API key is required. Set it via --api-key or OPENAI_API_KEY environment variable.")

    transcribe_slides(
        args.pptx_file,
        args.output_dir,
        model=args.model,
        language=args.language,
        task=args.task,
        prefix=args.prefix,
        api_key=args.api_key
    )


if __name__ == "__main__":
    main()
