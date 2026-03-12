import argparse
from pathlib import Path

try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None

if load_dotenv:
    load_dotenv()

from .converter import convert_file
from .errors import ConversionError


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert tests to MyTest format.")
    parser.add_argument("inputs", nargs="+", help="Input files (pdf, docx, xlsx, txt).")
    parser.add_argument("-o", "--output", help="Output file (only for single input).")
    parser.add_argument("--keep-numbers", action="store_true", help="Keep original numbers.")
    args = parser.parse_args()

    if args.output and len(args.inputs) > 1:
        raise SystemExit("--output can be used only with a single input file.")

    for input_path in args.inputs:
        try:
            output_text = convert_file(input_path, keep_numbers=args.keep_numbers)
        except ConversionError as exc:
            print(f"Error: {input_path}: {exc}")
            continue

        if args.output:
            out_path = Path(args.output)
        else:
            out_path = Path(input_path).with_suffix(".mytest.txt")
        out_path.write_text(output_text, encoding="utf-8")
        print(f"Wrote {out_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
