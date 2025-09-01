import os
import re
from config import api_key
import pytest
import pandas as pd
from docx import Document
import openai

# Configure OpenAI key
openai.api_key = api_key

DOCX_FILE = "Test_Requirements_Document_LinkedIn.docx"
OUTPUT_FILE = "linkedin_testcases.xlsx"


def extract_requirements(docx_file):
    """Extract requirements (TR-, BR-, NFR-) from a docx file."""
    doc = Document(docx_file)
    requirements = []

    # Regex to match requirement IDs at the start of the line
    # Examples: "TR-1", "BR-2", "NFR-3"
    req_pattern = re.compile(r"^(TR|BR|NFR)-\d+", re.IGNORECASE)

    for para in doc.paragraphs:
        text = para.text.strip()

        # Remove leading bullets, dashes, asterisks, and whitespace
        text = re.sub(r"^[•\-\*\s]+", "", text)

        if not text:
            continue  # skip empty lines

        # Replace full-width colon with normal colon
        text = text.replace("：", ":")

        # Split on the first colon
        parts = text.split(":", 1)
        if len(parts) != 2:
            continue  # skip lines without a colon

        req_id, req_text = parts
        req_id = req_id.strip()
        req_text = req_text.strip()

        # Check if the ID matches TR-/BR-/NFR- pattern
        if req_pattern.match(req_id):
            requirements.append({"id": req_id, "text": req_text})
    return requirements


def generate_testcases(requirement_id, requirement_text):
    """Call OpenAI API to generate test cases for a requirement"""
    prompt = f"""
    You are a QA engineer. Generate 3 manual test cases for this requirement.
    Return as a CSV-like table with columns:
    Test Case ID, Requirement ID, Test Description, Preconditions, Test Steps, Expected Result, Priority.

    Requirement ({requirement_id}): {requirement_text}
    """

    response = openai.ChatCompletion.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2
    )

    return response.choices[0].message["content"]


def parse_testcases(raw_text, requirement_id):
    """Parse AI-generated test cases from markdown or plain text into structured dicts"""
    rows = []
    lines = raw_text.splitlines()

    for line in lines:
        line = line.strip()
        if not line or line.startswith(("-", "Test Case ID")):
            continue

        # Check for markdown table row
        if line.startswith("|"):
            # Split on '|' and strip spaces
            parts = [p.strip() for p in line.strip("|").split("|")]
            if len(parts) >= 7:
                rows.append({
                    "Test Case ID": parts[0],
                    "Requirement ID": parts[1],
                    "Test Description": parts[2],
                    "Preconditions": parts[3],
                    "Test Steps": parts[4],
                    "Expected Result": parts[5],
                    "Priority": parts[6] if parts[6] else "Medium"
                })
    return rows


@pytest.mark.parametrize("docx_file", [DOCX_FILE])
def test_generate_testcases_from_requirements(docx_file):
    """Test that requirements are converted into test cases and exported to Excel"""
    requirements = extract_requirements(docx_file)
    assert requirements, "No requirements found in document!"

    all_testcases = []
    for req in requirements[:3]:  # limit to 3 for test speed
        raw_output = generate_testcases(req["id"], req["text"])
        testcases = parse_testcases(raw_output, req["id"])
        all_testcases.extend(testcases)

    assert all_testcases, "No test cases generated from requirements!"

    # Export to Excel
    df = pd.DataFrame(all_testcases)
    df.to_excel(OUTPUT_FILE, index=False)

    # Validate Excel file
    assert os.path.exists(OUTPUT_FILE), "Excel file not created!"
    exported = pd.read_excel(OUTPUT_FILE)
    assert not exported.empty, "Exported Excel is empty!"
