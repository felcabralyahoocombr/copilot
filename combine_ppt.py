"""Combine slides from multiple PowerPoint files into one presentation.

Install dependencies with:
    python -m pip install python-pptx pptxcompose

Example:
    python combine_ppt.py part1.pptx part2.pptx -o combined.pptx
"""

import argparse
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Combine the slides from multiple PPTX files."
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        type=Path,
        help="PPTX files to combine, in the order they should appear",
    )
    parser.add_argument(
        "-o",
        "--output",
        required=True,
        type=Path,
        help="path for the combined PPTX file",
    )
    args = parser.parse_args()

    if len(args.inputs) < 2:
        parser.error("provide at least two input PPTX files")

    output_path = args.output.resolve()
    for input_path in args.inputs:
        if not input_path.is_file():
            parser.error(f"input file does not exist: {input_path}")
        if input_path.resolve() == output_path:
            parser.error(f"output must not overwrite an input file: {input_path}")

    from pptx import Presentation
    from pptxcompose import Composer

    composer = Composer(Presentation(str(args.inputs[0])))
    for input_path in args.inputs[1:]:
        composer.append(Presentation(str(input_path)))

    composer.save(str(args.output))
    print(f"Combined {len(args.inputs)} presentations into {args.output}")


if __name__ == "__main__":
    main()