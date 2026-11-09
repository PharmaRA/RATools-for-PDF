"""Explicit CLI for version 1 qualification manifests; no GUI required."""

import argparse
from pathlib import Path
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from ratools_pdf.services.qualification import run_qualification


def main(argv=None):
    parser = argparse.ArgumentParser(description="Read-only PDF qualification evidence")
    parser.add_argument("--manifest", required=True, help="Versioned JSON manifest")
    parser.add_argument("--output", required=True, help="New evidence directory (must not exist)")
    parser.add_argument("--samples-root", help="Explicit sample root override")
    parser.add_argument("--skip-qpdf", action="store_true", help="Leave qpdf checks pending")
    args = parser.parse_args(argv)
    try:
        report = run_qualification(args.manifest, args.output,
            samples_root=args.samples_root, qpdf_enabled=not args.skip_qpdf)
    except (ValueError, OSError) as exc:
        print(f"Qualification could not start: {type(exc).__name__}", file=sys.stderr)
        return 2
    print(f"Qualification status: {report['status']}; sample counts: {report['summary']}")
    print("Evidence file: qualification-report.json")
    return 0 if report["status"] == "passed" else 1 if report["status"] == "failed" else 2


if __name__ == "__main__":
    raise SystemExit(main())
