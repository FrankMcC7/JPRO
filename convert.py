import glob
import os

# Configuration: hard-coded prefix and output path
PREFIX = "6gb"  # Use the same prefix used during splitting
OUTFILE = r"C:\Users\abc"


def join_file(prefix: str, outfile: str) -> None:
    """
    Joins all files named '{prefix}.part###' into a single ZIP at 'outfile'.
    """
    parts = sorted(glob.glob(f"{prefix}.part*"))
    if not parts:
        print(f"No parts found with prefix '{prefix}'. Ensure files are in the current directory.")
        return

    with open(outfile, 'wb') as dst:
        for part in parts:
            with open(part, 'rb') as src:
                data = src.read()
                dst.write(data)
            print(f"Merged: {part} ({os.path.getsize(part)} bytes)")

    print(f"Reassembled ZIP created at '{outfile}'. ({os.path.getsize(outfile)} bytes)")

if __name__ == '__main__':
    print(f"Starting reassembly of '{PREFIX}.part*' into '{OUTFILE}'...")
    join_file(PREFIX, OUTFILE)
