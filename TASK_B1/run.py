"""
Entry-point script for the Vehicle Detection, Tracking, and Trajectory Extraction pipeline.

Usage:
    python run.py

Before running, ensure:
  1. You have placed your input video at the path specified in config.py.
  2. You have adjusted START_FRAME, END_FRAME, and other settings in config.py.
  3. All dependencies listed in requirements.txt are installed.
"""

import sys
import time
from tracker import run_pipeline


def main():
    print("=" * 70)
    print("  Vehicle Detection, Tracking, and Trajectory Extraction Pipeline")
    print("=" * 70)
    print()

    start_time = time.time()

    try:
        video_path, csv_path = run_pipeline()
    except FileNotFoundError as exc:
        print(f"\n[ERROR] {exc}")
        print("        Please set INPUT_VIDEO in config.py to a valid video file.")
        sys.exit(1)
    except Exception as exc:
        print(f"\n[ERROR] Pipeline failed: {exc}")
        raise

    elapsed = time.time() - start_time
    minutes, seconds = divmod(elapsed, 60)

    print()
    print("-" * 70)
    print("  Pipeline completed successfully.")
    print(f"  Elapsed time    : {int(minutes)}m {seconds:.1f}s")
    print(f"  Annotated video : {video_path}")
    print(f"  Trajectory CSV  : {csv_path}")
    print("-" * 70)


if __name__ == "__main__":
    main()
