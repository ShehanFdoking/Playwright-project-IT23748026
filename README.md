# Playwright-project-IT23748026

A test automation framework using Playwright and Python to automate browser-based testing for web applications.

## Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

## Installation

### 1. Clone or Set Up the Project

Navigate to the project directory:

```bash
cd d:\test_automation_IT23748026
```

### 2. Create and Activate Virtual Environment

Create a Python virtual environment:

```bash
python -m venv .venv
```

Activate the virtual environment:

**On Windows (PowerShell):**

```bash
.\.venv\Scripts\Activate.ps1
```

**On Windows (Command Prompt):**

```bash
.venv\Scripts\activate.bat
```

**On macOS/Linux:**

```bash
source .venv/bin/activate
```

### 3. Install Dependencies

Install the required packages:

```bash
pip install playwright openpyxl
```

Install Playwright browsers (required for automation):

```bash
playwright install
```

## Running Tests

### Basic Test Run

Run the test automation script with default settings:

```bash
python test_automation.py
```

This will:

- Read test cases from the Excel file
- Execute tests against the default frontend URL (https://www.pixelssuite.com/chat-translator)
- Populate the Excel file with actual output and test results

### Advanced Options

The test script supports various command-line arguments for customization:

```bash
python test_automation.py [options]
```

**Common Options:**

- `--excel <path>` - Path to Excel file containing test cases (default: auto-detected)
- `--sheet <name>` - Excel sheet name (default: " Test cases")
- `--url <url>` - Frontend URL to test (default: https://www.pixelssuite.com/chat-translator)
- `--output <path>` - Output file path for results (default: overwrites input Excel file)
- `--headless` - Run browser in headless mode (no UI)
- `--wait-ms <ms>` - Wait time for elements (default: 5000ms)
- `--retries <n>` - Number of retries for failed tests (default: 8)
- `--timeout-ms <ms>` - Timeout for operations (default: 60000ms)
- `--keep-open` - Keep browser open after tests complete

**Example - Run tests with custom URL and save output to new file:**

```bash
python test_automation.py --url "https://example.com" --output results.xlsx --headless
```

**Example - Run tests with more retries:**

```bash
python test_automation.py --retries 10
```

## Project Structure

- `test_automation.py` - Main test automation script
- `fill_test_cases.py` - Utility to populate test case data
- `Assignment 1 - Test cases_IT23748026.xlsx` - Excel file containing test cases and results
- `.venv/` - Python virtual environment directory

## Deactivating Virtual Environment

When finished, deactivate the virtual environment:

```bash
deactivate
```
