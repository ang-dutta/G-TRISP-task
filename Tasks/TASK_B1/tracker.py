"""
Vehicle Detection, Tracking, and Trajectory Extraction Pipeline.

This module contains the core logic for:
  1. Loading a YOLOv8 pretrained model.
  2. Reading frames from the input video within the specified frame range.
  3. Running detection + tracking (ByteTrack) on each frame.
  4. Collecting per-frame trajectory data (bounding boxes, centroids, track IDs).
  5. Writing an annotated output video with bounding boxes, labels, and trails.
  6. Saving a structured CSV with all trajectory information.
"""

import os
import cv2
import numpy as np
import pandas as pd
from collections import defaultdict
from ultralytics import YOLO

import config


def load_model():
    """Load the YOLOv8 model specified in config."""
    print(f"[INFO] Loading model: {config.MODEL_VARIANT}")
    model = YOLO(config.MODEL_VARIANT)
    return model


def open_video(video_path):
    """Open a video file and return the capture object along with metadata."""
    if not os.path.isfile(video_path):
        raise FileNotFoundError(f"Input video not found: {video_path}")

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        raise RuntimeError(f"Could not open video: {video_path}")

    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    fps = cap.get(cv2.CAP_PROP_FPS)
    width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))

    print(f"[INFO] Video opened: {video_path}")
    print(f"       Resolution : {width}x{height}")
    print(f"       FPS        : {fps}")
    print(f"       Total Frames: {total_frames}")

    return cap, total_frames, fps, width, height


def resolve_frame_range(total_frames):
    """Compute the effective start and end frame indices."""
    start = config.START_FRAME if config.START_FRAME is not None else 0
    end = config.END_FRAME if config.END_FRAME is not None else total_frames

    start = max(0, start)
    end = min(end, total_frames)

    if start >= end:
        raise ValueError(
            f"Invalid frame range: START_FRAME={start}, END_FRAME={end}. "
            f"Video has {total_frames} frames."
        )

    print(f"[INFO] Processing frames {start} to {end - 1} "
          f"({end - start} frames total)")
    return start, end


def get_color_for_id(track_id):
    """Generate a deterministic color for a given track ID."""
    np.random.seed(int(track_id) * 7 + 13)
    color = tuple(int(c) for c in np.random.randint(80, 255, size=3))
    return color


def draw_annotations(frame, boxes, track_ids, class_ids, trails):
    """
    Draw bounding boxes, labels, and optional trajectory trails on a frame.

    Parameters
    ----------
    frame : np.ndarray
        The video frame to annotate (modified in place).
    boxes : np.ndarray
        Array of shape (N, 4) with xyxy bounding box coordinates.
    track_ids : np.ndarray
        Array of shape (N,) with integer track IDs.
    class_ids : np.ndarray
        Array of shape (N,) with integer class IDs.
    trails : dict
        Mapping from track_id to list of (cx, cy) centroid positions.
    """
    for box, tid, cid in zip(boxes, track_ids, class_ids):
        x1, y1, x2, y2 = map(int, box)
        tid = int(tid)
        cid = int(cid)
        color = get_color_for_id(tid)
        class_label = config.VEHICLE_CLASS_NAMES.get(cid, f"class_{cid}")
        label = f"ID:{tid} {class_label}"

        # Bounding box
        cv2.rectangle(frame, (x1, y1), (x2, y2), color, config.BOX_THICKNESS)

        # Label background
        (tw, th), _ = cv2.getTextSize(
            label, cv2.FONT_HERSHEY_SIMPLEX, config.FONT_SCALE, 1
        )
        cv2.rectangle(frame, (x1, y1 - th - 8), (x1 + tw + 4, y1), color, -1)
        cv2.putText(
            frame, label, (x1 + 2, y1 - 4),
            cv2.FONT_HERSHEY_SIMPLEX, config.FONT_SCALE,
            (255, 255, 255), 1, cv2.LINE_AA
        )

        # Centroid dot
        cx = (x1 + x2) // 2
        cy = (y1 + y2) // 2
        cv2.circle(frame, (cx, cy), 4, color, -1)

        # Trajectory trail
        if config.DRAW_TRAJECTORY_TRAIL and tid in trails:
            points = trails[tid]
            for i in range(1, len(points)):
                thickness = max(1, int(2 * (i / len(points))))
                cv2.line(frame, points[i - 1], points[i], color, thickness,
                         cv2.LINE_AA)


def run_pipeline():
    """
    Execute the full detection, tracking, and trajectory extraction pipeline.

    Returns the path to the annotated video and the trajectory CSV.
    """
    # Prepare output directory
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)

    # Load model
    model = load_model()

    # Open video
    cap, total_frames, fps, width, height = open_video(config.INPUT_VIDEO)

    # Determine frame range
    start_frame, end_frame = resolve_frame_range(total_frames)

    # Seek to start frame
    cap.set(cv2.CAP_PROP_POS_FRAMES, start_frame)

    # Prepare video writer
    output_video_path = os.path.join(
        config.OUTPUT_DIR, config.ANNOTATED_VIDEO_FILENAME
    )
    fourcc = cv2.VideoWriter_fourcc(*"mp4v")
    writer = cv2.VideoWriter(output_video_path, fourcc, fps, (width, height))

    # Data collection
    trajectory_records = []
    trails = defaultdict(list)  # track_id -> list of (cx, cy)

    num_frames = end_frame - start_frame
    print(f"[INFO] Starting processing ...")

    for idx in range(num_frames):
        ret, frame = cap.read()
        if not ret:
            print(f"[WARN] Could not read frame at index {start_frame + idx}. "
                  f"Stopping early.")
            break

        current_frame = start_frame + idx

        # Run detection + tracking
        results = model.track(
            source=frame,
            tracker=config.TRACKER_CONFIG,
            conf=config.CONFIDENCE_THRESHOLD,
            iou=config.IOU_THRESHOLD,
            imgsz=config.IMGSZ,
            classes=config.VEHICLE_CLASS_IDS,
            persist=True,
            verbose=False,
        )

        result = results[0]

        if result.boxes is not None and result.boxes.id is not None:
            boxes = result.boxes.xyxy.cpu().numpy()
            track_ids = result.boxes.id.cpu().numpy().astype(int)
            class_ids = result.boxes.cls.cpu().numpy().astype(int)
            confidences = result.boxes.conf.cpu().numpy()

            # Stationary filter: skip tracks whose centroid has barely moved.
            def is_stationary(tid):
                if not config.ENABLE_STATIONARY_FILTER:
                    return False
                history = trails[int(tid)]
                if len(history) < config.STATIONARY_WINDOW:
                    return False
                recent = history[-config.STATIONARY_WINDOW:]
                xs = [p[0] for p in recent]
                ys = [p[1] for p in recent]
                displacement = ((max(xs) - min(xs)) ** 2 +
                                (max(ys) - min(ys)) ** 2) ** 0.5
                return displacement < config.STATIONARY_PIXEL_THRESHOLD

            # Collect trajectory data (moving tracks only)
            for box, tid, cid, conf in zip(boxes, track_ids, class_ids,
                                           confidences):
                x1, y1, x2, y2 = box
                cx = (x1 + x2) / 2.0
                cy = (y1 + y2) / 2.0

                # Update trails first so is_stationary has data to work with.
                pt = (int(cx), int(cy))
                trails[int(tid)].append(pt)
                if len(trails[int(tid)]) > config.TRAIL_LENGTH:
                    trails[int(tid)] = trails[int(tid)][-config.TRAIL_LENGTH:]

                if is_stationary(int(tid)):
                    continue

                class_label = config.VEHICLE_CLASS_NAMES.get(
                    int(cid), f"class_{int(cid)}"
                )

                trajectory_records.append({
                    "frame_number": current_frame,
                    "track_id": int(tid),
                    "vehicle_class": class_label,
                    "x1": round(float(x1), 2),
                    "y1": round(float(y1), 2),
                    "x2": round(float(x2), 2),
                    "y2": round(float(y2), 2),
                    "centroid_x": round(float(cx), 2),
                    "centroid_y": round(float(cy), 2),
                    "confidence": round(float(conf), 4),
                })

            # Build filtered arrays for annotation (exclude stationary tracks).
            keep = [i for i, tid in enumerate(track_ids)
                    if not is_stationary(int(tid))]
            if keep:
                draw_annotations(
                    frame,
                    boxes[keep],
                    track_ids[keep],
                    class_ids[keep],
                    trails,
                )

        # Write annotated frame
        writer.write(frame)

        # Progress reporting
        if (idx + 1) % 100 == 0 or (idx + 1) == num_frames:
            pct = (idx + 1) / num_frames * 100
            print(f"[INFO] Processed {idx + 1}/{num_frames} frames ({pct:.1f}%)")

    # Cleanup
    cap.release()
    writer.release()
    print(f"[INFO] Annotated video saved to: {output_video_path}")

    # Save trajectory CSV
    csv_path = os.path.join(config.OUTPUT_DIR, config.TRAJECTORY_CSV_FILENAME)
    if trajectory_records:
        df = pd.DataFrame(trajectory_records)
        df.to_csv(csv_path, index=False)
        unique_tracks = df["track_id"].nunique()
        print(f"[INFO] Trajectory CSV saved to: {csv_path}")
        print(f"       Total detections : {len(df)}")
        print(f"       Unique track IDs : {unique_tracks}")
    else:
        # Write header-only CSV so downstream tools do not break.
        pd.DataFrame(columns=[
            "frame_number", "track_id", "vehicle_class",
            "x1", "y1", "x2", "y2",
            "centroid_x", "centroid_y", "confidence"
        ]).to_csv(csv_path, index=False)
        print(f"[WARN] No vehicle detections found. Empty CSV saved to: {csv_path}")

    return output_video_path, csv_path
