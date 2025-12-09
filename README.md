# Barchart Screenshot Tool

A Python tool to capture screenshots of the Barchart put-call ratios page using Playwright.

## Features

- Captures full-page screenshots of Barchart financial data pages
- Saves screenshots with timestamps for easy tracking
- Uses headless browser automation (no visible browser window)

## Prerequisites

- Python 3.10 or higher
- pip (Python package manager)

## Installation

1. Install the required Python packages:
   ```bash
   pip install -r requirements.txt
   ```

2. Install Playwright browsers:
   ```bash
   playwright install chromium
   ```

## Usage

Run the screenshot script:

```bash
python screenshot.py
```

Screenshots will be saved to the `screenshots/` directory with timestamps.

## Configuration

You can customize the screenshot by modifying the `take_screenshot()` function parameters:

- `url`: Change the target URL
- `output_dir`: Change the output directory
- `filename`: Specify a custom filename

## Output

Screenshots are saved as PNG files in the `screenshots/` directory with the naming format:
`barchart_put_call_ratios_YYYYMMDD_HHMMSS.png`
