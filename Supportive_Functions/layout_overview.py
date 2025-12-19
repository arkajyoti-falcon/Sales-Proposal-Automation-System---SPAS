import cv2
import numpy as np
import torch
from PIL import Image
from transformers import AutoProcessor, AutoModelForZeroShotObjectDetection

# --------------- CONFIG ---------------
IMAGE_PATH = "layout.png"   # your layout image

LEGENDS = [
    "loading conveyor zone 1",
    "highway lines",
    "induct zone",
    "ICR tunnel",
    "linear cross belt sorter",
    "live chutes",
    "collection chutes",
    "recirculation line",
    "RC line",
]

BOX_THRESHOLD = 0.25
TEXT_THRESHOLD = 0.25

MODEL_ID = "IDEA-Research/grounding-dino-tiny"
# --------------------------------------


def get_device():
    try:
        from accelerate import Accelerator
        return Accelerator().device
    except Exception:
        return torch.device("cuda" if torch.cuda.is_available() else "cpu")


def main():
    device = get_device()
    print(f"[INFO] Using device: {device}")

    # ---- Load model & processor ----
    print("[INFO] Loading model...")
    processor = AutoProcessor.from_pretrained(MODEL_ID)
    model = AutoModelForZeroShotObjectDetection.from_pretrained(MODEL_ID).to(device)

    # ---- Load image ----
    image_pil = Image.open(IMAGE_PATH).convert("RGB")
    W, H = image_pil.size
    print(f"[INFO] Image size: {W} x {H}")

    # ---- Prepare inputs ----
    # GroundingDINO expects list-of-list for text: [ [label1, label2, ...] ]
    text_labels = [LEGENDS]

    inputs = processor(
        images=image_pil,
        text=text_labels,
        return_tensors="pt"
    ).to(device)

    with torch.no_grad():
        outputs = model(**inputs)

    # HF post-process: uses `threshold` (not box_threshold)
    results = processor.post_process_grounded_object_detection(
        outputs=outputs,
        input_ids=inputs.input_ids,
        threshold=BOX_THRESHOLD,
        text_threshold=TEXT_THRESHOLD,
        target_sizes=[(H, W)],  # (height, width)
    )

    result = results[0]
    boxes = result["boxes"]      # (N, 4) absolute XYXY (pixels)
    scores = result["scores"]    # (N,)
    labels = result["labels"]    # (N,) strings corresponding to text prompts

    # ---- Pick best box per legend string ----
    # key: legend_name (string) -> (score, box_tensor)
    best_by_label = {}

    for box, score, label in zip(boxes, scores, labels):
        legend_name = str(label)
        s = float(score.item())

        if s < BOX_THRESHOLD:
            continue

        # ignore any labels that are not exactly in our legend list
        if legend_name not in LEGENDS:
            continue

        if legend_name not in best_by_label or s > best_by_label[legend_name][0]:
            best_by_label[legend_name] = (s, box)

    # ---- Draw annotations ----
    annotated = cv2.cvtColor(np.array(image_pil), cv2.COLOR_RGB2BGR)

    colors = [
        (255, 0, 0),
        (0, 255, 0),
        (0, 0, 255),
        (255, 255, 0),
        (255, 0, 255),
        (0, 255, 255),
        (128, 0, 255),
        (0, 128, 255),
        (255, 128, 0),
    ]

    all_labels = []
    display_idx = 1

    # iterate in the same order as LEGENDS for numbering
    for legend_name in LEGENDS:
        if legend_name not in best_by_label:
            print(f"[WARN] No box found for: {legend_name}")
            continue

        score, box = best_by_label[legend_name]
        print(f"[INFO] {legend_name}: best score={score:.3f}")

        # box is already [x_min, y_min, x_max, y_max] in pixels
        x1, y1, x2, y2 = [int(v) for v in box.tolist()]

        color = colors[(display_idx - 1) % len(colors)]
        cv2.rectangle(annotated, (x1, y1), (x2, y2), color, 2)

        # Number bubble near top-left of the box
        label_text = str(display_idx)
        (tw, th), _ = cv2.getTextSize(
            label_text, cv2.FONT_HERSHEY_SIMPLEX, 0.7, 2
        )
        y_top = max(0, y1 - th - 10)
        cv2.rectangle(
            annotated,
            (x1, y_top),
            (x1 + tw + 10, y1),
            color,
            -1,
        )
        cv2.putText(
            annotated,
            label_text,
            (x1 + 5, max(0, y1 - 5)),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.7,
            (0, 0, 0),
            2,
        )

        all_labels.append((display_idx, legend_name))
        display_idx += 1

    # ---- Draw legend box ----
    legend_img = annotated.copy()
    x0, y0 = 20, 40
    line_h = 25

    if all_labels:
        box_height = line_h * len(all_labels) + 20
    else:
        box_height = 40

    cv2.rectangle(
        legend_img,
        (x0 - 10, y0 - 30),
        (x0 + 450, y0 + box_height),
        (255, 255, 255),
        -1,
    )
    cv2.rectangle(
        legend_img,
        (x0 - 10, y0 - 30),
        (x0 + 450, y0 + box_height),
        (0, 0, 0),
        1,
    )

    cv2.putText(
        legend_img,
        "Legend",
        (x0, y0 - 10),
        cv2.FONT_HERSHEY_SIMPLEX,
        0.7,
        (0, 0, 0),
        2,
    )

    for i, (idx_num, name) in enumerate(all_labels):
        text = f"{idx_num} - {name}"
        cv2.putText(
            legend_img,
            text,
            (x0, y0 + i * line_h),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.6,
            (0, 0, 0),
            1,
        )

    out_path = "layout_annotated.png"
    cv2.imwrite(out_path, legend_img)
    print(f"[INFO] Saved: {out_path}")


if __name__ == "__main__":
    main()
