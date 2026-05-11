# Vehicle Detection, Tracking, and Trajectory Extraction

## Overview

This project implements a vehicle detection, tracking, and trajectory extraction pipeline for traffic video analysis. It uses a pretrained YOLOv8 model for detection and ByteTrack for multi-object tracking. The system produces an annotated video with bounding boxes and trajectory trails, along with a structured CSV file containing per-frame tracking data.

---

## Model Used

| Component  | Details                                                                 |
|------------|-------------------------------------------------------------------------|
| Detector   | YOLOv8n (nano variant, pretrained on COCO -- 80 classes)                |
| Tracker    | ByteTrack (association-based tracker using Kalman filtering)            |
| Framework  | Ultralytics (`ultralytics` Python package)                              |
| Classes    | car (2), motorcycle (3), bus (5), truck (7) from the COCO class set     |

YOLOv8n was chosen for its balance between inference speed and accuracy. For higher accuracy on dense traffic scenes, swap to `yolov8s.pt`, `yolov8m.pt`, or `yolov8x.pt` in `config.py`.

---

## Tracking Method

ByteTrack is the default tracker provided by the Ultralytics library. It works in two stages:

1. **High-confidence matching** -- Associates high-confidence detections with existing tracks using IoU-based matching.
2. **Low-confidence recovery** -- Unmatched tracks are re-associated with lower-confidence detections, reducing ID switches caused by occlusion or brief disappearances.

The tracker uses a Kalman filter to predict each track's next position, enabling robust association even when detections are momentarily lost.

---

## Output Format

### Annotated Video (`output/annotated_output.mp4`)

Each frame in the output video contains:
- Colored bounding boxes around detected vehicles.
- Text labels showing the track ID and vehicle class.
- Centroid dots at the center of each bounding box.
- Trajectory trails showing recent centroid positions per track.

### Trajectory CSV (`output/trajectories.csv`)

| Column         | Type    | Description                                      |
|----------------|---------|--------------------------------------------------|
| frame_number   | int     | Frame index in the original video                |
| track_id       | int     | Unique ID assigned to the tracked vehicle        |
| vehicle_class  | string  | Vehicle type (car, motorcycle, bus, truck)        |
| x1             | float   | Bounding box top-left x coordinate               |
| y1             | float   | Bounding box top-left y coordinate               |
| x2             | float   | Bounding box bottom-right x coordinate           |
| y2             | float   | Bounding box bottom-right y coordinate           |
| centroid_x     | float   | Centroid x coordinate ((x1+x2)/2)                |
| centroid_y     | float   | Centroid y coordinate ((y1+y2)/2)                |
| confidence     | float   | Detection confidence score (0 to 1)              |

---

## Project Structure

```
TASK_B2/
|-- config.py                  # All configurable parameters
|-- tracker.py                 # Core detection, tracking, and annotation logic
|-- run.py                     # CLI entry point
|-- vehicle_tracking.ipynb     # Jupyter notebook (interactive version)
|-- requirements.txt           # Python dependencies
|-- README.md                  # This file
|-- input/                     # Place your input video here
|   +-- traffic_video.mp4
+-- output/                    # Generated outputs (created automatically)
    |-- annotated_output.mp4
    +-- trajectories.csv
```

---

## Steps to Run

### Prerequisites

- Python 3.8 or higher
- A traffic video file (MP4 or AVI format)

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Place Your Input Video

Copy your traffic video into the `input/` directory and rename it to `traffic_video.mp4`, or update the `INPUT_VIDEO` path in `config.py`.

### 3. Configure Frame Range (Optional)

Open `config.py` and set `START_FRAME` and `END_FRAME` to process only a portion of the video:

```python
START_FRAME = 0       # First frame to process
END_FRAME = 500       # Last frame (exclusive). Set to None for full video.
```

### 4a. Run via Script

```bash
python run.py
```

### 4b. Run via Notebook

Open `vehicle_tracking.ipynb` in Jupyter Notebook or JupyterLab and execute cells sequentially. The notebook includes the same pipeline with inline configuration and trajectory visualization.

### 5. View Outputs

After execution completes, find the results in the `output/` directory:
- `annotated_output.mp4` -- Video with bounding boxes, track IDs, and trails.
- `trajectories.csv` -- Structured CSV with all tracking data.

---

## Configuration Reference

All parameters are located in `config.py`:

| Parameter                | Default              | Description                                       |
|--------------------------|----------------------|---------------------------------------------------|
| INPUT_VIDEO              | input/traffic_video.mp4 | Path to input video                            |
| OUTPUT_DIR               | output               | Output directory for results                      |
| START_FRAME              | 0                    | First frame to process                            |
| END_FRAME                | None (full video)    | Last frame to process (exclusive)                 |
| MODEL_VARIANT            | yolov8n.pt           | YOLOv8 model variant                              |
| CONFIDENCE_THRESHOLD     | 0.35                 | Minimum detection confidence                      |
| IOU_THRESHOLD            | 0.5                  | NMS IoU threshold                                 |
| TRACKER_CONFIG           | bytetrack.yaml       | Tracker algorithm config                          |
| VEHICLE_CLASS_IDS        | [2, 3, 5, 7]        | COCO class IDs for vehicles                       |
| BOX_THICKNESS            | 2                    | Bounding box line thickness                       |
| FONT_SCALE               | 0.6                  | Label text size                                   |
| DRAW_TRAJECTORY_TRAIL    | True                 | Enable/disable centroid trail drawing              |
| TRAIL_LENGTH             | 50                   | Max trail points per track                        |

---

## Notes

- The YOLOv8 model weights are downloaded automatically on first run.
- GPU acceleration is used automatically if a CUDA-compatible GPU is available. Otherwise, the pipeline runs on CPU.
- The `bytetrack.yaml` configuration file is bundled with the `ultralytics` package and does not need to be created manually.
- For very long videos, consider processing in segments by adjusting START_FRAME and END_FRAME.
