import argparse
import struct
from pathlib import Path

import pefile


WINDOWS_GUI_SUBSYSTEM = 2


def patch_subsystem(exe_path: Path, subsystem: int) -> None:
    pe = pefile.PE(str(exe_path))
    offset = pe.OPTIONAL_HEADER.get_field_absolute_offset("Subsystem")
    pe.close()

    data = bytearray(exe_path.read_bytes())
    data[offset:offset + 2] = struct.pack("<H", subsystem)
    exe_path.write_bytes(data)


def main() -> int:
    parser = argparse.ArgumentParser(description="Patch PE subsystem field.")
    parser.add_argument("exe_path", help="Path to the target exe")
    parser.add_argument("--windows-gui", action="store_true", help="Set subsystem to Windows GUI")
    args = parser.parse_args()

    subsystem = WINDOWS_GUI_SUBSYSTEM if args.windows_gui else None
    if subsystem is None:
        parser.error("No subsystem patch option provided")

    patch_subsystem(Path(args.exe_path), subsystem)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())