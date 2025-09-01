# Testcase Generator

A Python script to automatically generate manual test cases from requirement documents (.docx) and export them to Excel (.xlsx), which can then be uploaded to test management tools like TestRail or others.

---

## Features

* Extracts functional, non-functional, and business requirements from `.docx` requirement documents.
* Generates multiple test cases per requirement using OpenAI GPT API.
* Exports test cases to an Excel file ready for upload.
* Includes a sample `.docx` requirement file for testing.

---

## Requirements

* Python 3.11
* pip (Python package manager)

### Python Packages

The script depends on the following Python packages:

* `openai`
* `python-docx`
* `pandas`
* `openpyxl`
* `pytest` (for testing)

All dependencies are listed in `requirements.txt`.

---

## Setup Instructions

### 1. Clone the Repository

```bash
git clone <your-repo-url>
cd <your-repo-folder>
```

### 2. Create a Virtual Environment (Recommended)

```bash
python3.11 -m venv .venv
source .venv/bin/activate    # On Windows use: .venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Configure OpenAI API Key

Create a `.env` file in the project root (or use your environment variables) with the following content:

```bash
OPENAI_API_KEY=your_openai_api_key_here
```

> **Note:** Make sure the API key has access to the models you intend to use.

---

## Usage

### Generate Test Cases

```bash
python test_case_generator.py --input_file path/to/requirements.docx --output_file linkedin_testcases.xlsx
```

* **`--input_file`**: Path to your requirements `.docx` file.
* **`--output_file`**: Path to save the generated `.xlsx` file. Default is `linkedin_testcases.xlsx`.

The resulting Excel file will contain test cases with the following columns:

* Test Case ID
* Requirement ID
* Test Description
* Preconditions
* Test Steps
* Expected Result
* Priority

This `.xlsx` can be used to upload test cases to test management tools like TestRail.

---

### Run Tests

To validate that the script is working properly, run the included pytest tests:

```bash
pytest test_case_generator.py
```

---

## Sample Requirement Document

A sample `.docx` file is included in the repository (`Test_Requirements_Document_LinkedIn.docx`).
You can replace this with your own requirement documents.

---

## Notes

* Ensure your environment has internet access for OpenAI API calls.
* The script currently limits test case generation to the first 3 requirements for testing. Modify in the script as needed.
* Make sure `openpyxl` is installed to enable Excel export.

---

## License

[MIT License](LICENSE)
