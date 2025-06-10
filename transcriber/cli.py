try:
    from gooey import Gooey
except Exception:  # pragma: no cover - Gooey is optional
    Gooey = None  # type: ignore

from .parser import build_parser
from .transcribe_slides import run

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
        run(args)
else:
    def main() -> None:  # pragma: no cover - only used when Gooey is installed
        parser = build_parser(gui=False)
        args = parser.parse_args()
        run(args)


if __name__ == "__main__":
    main()
