"""
Configuration file for the Vehicle Detection, Tracking, and Trajectory Extraction pipeline.

All user-configurable parameters are defined here. Modify these values
before running the pipeline to suit your input video and requirements.
"""

import os

# ---------------------------------------------------------------------------
# Path Configuration
# ---------------------------------------------------------------------------

# Path to the input traffic video file.
INPUT_VIDEO = os.path.join(os.path.dirname(__file__), "input", "traffic_video.mp4")

# Directory where all outputs (annotated video, CSV, logs) will be saved.
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")

# ---------------------------------------------------------------------------
# Frame Range
# ---------------------------------------------------------------------------

# Set START_FRAME and END_FRAME to process only a portion of the video.
# Use None to start from the beginning or process until the end.
START_FRAME = 0
END_FRAME = 500  # Set to an integer (e.g., 500) to stop at that frame.

# ---------------------------------------------------------------------------
# Model Configuration
# ---------------------------------------------------------------------------

# YOLOv8 model variant. Options: yolov8n, yolov8s, yolov8m, yolov8l, yolov8x
# yolov8s gives significantly better class discrimination than yolov8n,
# particularly for motorcycles and large vehicles, at modest extra cost.
MODEL_VARIANT = "yolov8s.pt"

# Detection confidence threshold. Detections below this score are discarded.
# Lowered to 0.25 to catch more 2-wheelers which typically score lower.
CONFIDENCE_THRESHOLD = 0.25

# IOU threshold for non-maximum suppression.
IOU_THRESHOLD = 0.45

# Inference image size (pixels). Higher values improve detection of small objects
# like motorcycles at the cost of speed. Must be a multiple of 32.
# Use 640 for speed, 1280 for accuracy on high-resolution footage.
IMGSZ = 1280

# ---------------------------------------------------------------------------
# Tracking Configuration
# ---------------------------------------------------------------------------

# Tracker configuration file shipped with ultralytics.
# Options: "bytetrack.yaml", "botsort.yaml"
TRACKER_CONFIG = "bytetrack.yaml"

# ---------------------------------------------------------------------------
# Vehicle Classes (COCO Dataset IDs)
# ---------------------------------------------------------------------------

# COCO class IDs that correspond to vehicle categories:
#   2 = car, 3 = motorcycle, 5 = bus, 7 = truck
VEHICLE_CLASS_IDS = [2, 3, 5, 7]

# Human-readable mapping from COCO class ID to vehicle type label.
VEHICLE_CLASS_NAMES = {
    2: "car",
    3: "motorcycle",
    5: "bus",
    7: "truck",
}

# ---------------------------------------------------------------------------
# Annotation / Visualization
# ---------------------------------------------------------------------------

# Thickness of bounding box lines drawn on the annotated video.
BOX_THICKNESS = 2

# Font scale for text labels on the annotated video.
FONT_SCALE = 0.6

# Whether to draw the trajectory trail (centroid path) on the annotated video.
DRAW_TRAJECTORY_TRAIL = True

# Maximum number of past centroid positions to keep per track for trail drawing.
TRAIL_LENGTH = 50

# ---------------------------------------------------------------------------
# Stationary Object Filter
# ---------------------------------------------------------------------------

# Enable filtering of tracks that do not move (e.g. bus shelters, signboards
# misclassified as trucks). A track is considered stationary and suppressed
# if its centroid displacement over the last STATIONARY_WINDOW frames is
# smaller than STATIONARY_PIXEL_THRESHOLD pixels.
ENABLE_STATIONARY_FILTER = True

# Number of recent frames to look back when measuring movement.
STATIONARY_WINDOW = 15

# Minimum pixel displacement required to be treated as a moving vehicle.
# Tracks that move less than this across STATIONARY_WINDOW frames are hidden.
STATIONARY_PIXEL_THRESHOLD = 8

# ---------------------------------------------------------------------------
# Output File Names
# ---------------------------------------------------------------------------

ANNOTATED_VIDEO_FILENAME = "annotated_output.mp4"
TRAJECTORY_CSV_FILENAME = "trajectories.csv"
