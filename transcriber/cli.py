import argparse
import os

try:
    from gooey import Gooey, GooeyParser
except Exception:  # pragma: no cover - Gooey is optional
    Gooey = None  # type: ignore
    GooeyParser = None  # type: ignore

from transcriber.transcribe_slides import transcribe_slides


def build_parser(gui: bool = False) -> argparse.ArgumentParser:
    """Create an argument parser.

    Parameters
    ----------
    gui: bool
        If ``True`` and Gooey is available, a :class:`GooeyParser` with widgets
        will be returned. Otherwise a standard :class:`argparse.ArgumentParser`
        is created.
    """
    if gui and GooeyParser:
        parser = GooeyParser(description="Transcribe audio from PowerPoint slides")
        io = parser.add_argument_group("Input/Output", gooey_options={"columns": 1})
        io.add_argument(
            "pptx",
            help="Path to the PowerPoint file",
            widget="FileChooser",
            gooey_options={"wildcard": "PowerPoint files (*.pptx)|*.pptx"},
        )
        io.add_argument("output", help="Directory to save transcriptions", widget="DirChooser")

        trans = parser.add_argument_group(
            "Transcription Settings", gooey_options={"columns": 1}
        )
        trans.add_argument(
            "--model",
            default="whisper-1",
            choices=["whisper-1"],
            help="OpenAI STT model to use",
            widget="Dropdown",
        )
        trans.add_argument(
            "--language",
            default="en",
            help="Language spoken in the audio (e.g., 'en' for English)",
        )
        trans.add_argument(
            "--task",
            choices=["transcribe", "translate"],
            default="transcribe",
            help="Transcription task",
            widget="Dropdown",
        )
        trans.add_argument("--prefix", default="slide", help="Prefix for output files")

        api = parser.add_argument_group("API Settings", gooey_options={"columns": 1})
        api.add_argument(
            "--api-key",
            default=os.environ.get("OPENAI_API_KEY", ""),
            help="OpenAI API key (or set OPENAI_API_KEY environment variable)",
            widget="PasswordField",
        )
    else:
        parser = argparse.ArgumentParser(
            description="Transcribe slide audio using the OpenAI Speech to Text API"
        )
        parser.add_argument("pptx", help="Path to PPTX file")
        parser.add_argument("output", help="Output directory for transcripts")
        parser.add_argument(
            "--model",
            default="whisper-1",
            choices=["whisper-1"],
            help="OpenAI STT model to use (currently only whisper-1 is available)",
        )
        parser.add_argument(
            "--language",
            default="en",
            help="Language spoken in the audio (e.g., 'en' for English)",
        )
        parser.add_argument(
            "--task",
            choices=["transcribe", "translate"],
            default="transcribe",
            help="Transcription task",
        )
        parser.add_argument("--prefix", default="slide", help="Prefix for output files")
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


if Gooey:

    @Gooey(
        program_name="Slides Transcriber",
        program_description="Extract audio from PowerPoint slides and transcribe them using OpenAI's Whisper model",
        default_size=(800, 600),
        tabbed_groups=True,
        navigation="TABBED",
    )
    def main() -> None:
        parser = build_parser(gui=True)
        args = parser.parse_args()
        if not args.api_key:
            parser.error(
                "OpenAI API key is required. Set it via --api-key or OPENAI_API_KEY environment variable."
            )
        transcribe_slides(
            args.pptx,
            args.output,
            model=args.model,
            language=args.language,
            task=args.task,
            prefix=args.prefix,
            api_key=args.api_key,
        )
else:

    def main() -> None:  # pragma: no cover - only used when Gooey is installed
        parser = build_parser(gui=False)
        args = parser.parse_args()
        if not args.api_key:
            parser.error(
                "OpenAI API key is required. Set it via --api-key or OPENAI_API_KEY environment variable."
            )
        transcribe_slides(
            args.pptx,
            args.output,
            model=args.model,
            language=args.language,
            task=args.task,
            prefix=args.prefix,
            api_key=args.api_key,
        )


if __name__ == "__main__":
    main()
