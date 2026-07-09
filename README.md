# Number OCR with Paddle

Extracts **phone numbers (10 digits)** and **SIM serial numbers (15 digits)** from SIM-card images using [PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR), and saves the results to an Excel file.

## Features

- Batch-processes a whole folder of images (`jpg`, `jpeg`, `png`, `bmp`, `tiff`, `tif` — case-insensitive)
- Optional recursive scan of subfolders (`-r`)
- Tolerates separators inside numbers (e.g. `0912 345-6789` is still detected)
- Excel report with auto-fitted columns + summary statistics
- One bad image never stops the batch (errors are logged per file)
- Runs on GPU or CPU (`--device gpu:0` / `--device cpu`)

## Installation

Create and activate a virtual environment:

```bash
python -m venv venv

# Linux / macOS
source venv/bin/activate

# Windows
venv\Scripts\activate
```

Install the dependencies:

```bash
pip install -r requirements.txt
```

> **Note:** `requirements.txt` installs the GPU build of PaddlePaddle (requires CUDA). On a CPU-only machine, install `paddlepaddle==2.6.2` instead of `paddlepaddle-gpu==2.6.2`, and run with `--device cpu`.
>
> You must be connected to the internet the first time you run the app so PaddleOCR can download its models.

## Usage

```bash
# Process images in the current folder
python main.py

# Process a specific folder and choose the output file
python main.py path/to/images -o results.xlsx

# Include subfolders, run on CPU
python main.py path/to/images -r --device cpu
```

All options:

| Option | Description | Default |
|---|---|---|
| `folder` | Folder containing the images | current folder |
| `-o`, `--output` | Output Excel file path | `ocr_results_<timestamp>.xlsx` |
| `--device` | Inference device (`gpu:0`, `cpu`, ...) | `gpu:0` |
| `--lang` | OCR language | `en` |
| `-r`, `--recursive` | Also search subfolders | off |

## Use as a library

```python
from main import process_images_from_folder

df = process_images_from_folder("path/to/images", "results.xlsx", device="cpu")
```

## Output

The generated Excel file contains one row per image:

| Column | Description |
|---|---|
| Filename | Image file name |
| Full Path | Absolute path to the image |
| Phone Number | Detected 10-digit number (empty if not found) |
| Serial Number | Detected 15-digit number (empty if not found) |
| Status | `Success`, `No text detected`, or the error message |
| Processing Date | Timestamp of processing |
