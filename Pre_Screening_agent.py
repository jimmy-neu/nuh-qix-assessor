import os
import json
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
from typing import List
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# 1. Define the Structured Output Schema for the Screener
class ScreeningCheck(BaseModel):
    criterion: str = Field(description="The Level-4 exclusion rule being evaluated.")
    violation_found: bool = Field(description="True if the project meets this exclusion criteria.")
    evidence_found: str = Field(description="Specific evidence from the PDF or JSON justifying the decision.")

class ScreeningResult(BaseModel):
    is_eligible: bool = Field(description="Overall eligibility. True if NO violations are found.")
    primary_violation: str = Field(description="The specific rule that failed, or 'None'.")
    detailed_audit: List[ScreeningCheck]

# 2. Pre-Screening Agent Implementation
def run_pre_screening(json_path: str, pdf_path: str):
    """
    Reads the extracted JSON and original PDF to perform a Level-4 eligibility audit.
    """
    client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))

    # Load the extracted JSON data
    with open(json_path, "r", encoding="utf-8") as f:
        extracted_data = f.read()

    # Load the raw PDF file for grounding/verification
    with open(pdf_path, "rb") as f:
        doc_data = f.read()

    # Level-4 Exclusionary Criteria based on Evaluation Guidelines 
    level_4_rules = """
    Your sole responsibility is to be a strict auditor, flagging rule violations and missing criteria. You are looking only for reasons to classify the project as Level 4.
    Task: Analyze the provided project PDF and identify if it meets ANY of the following strict criteria for a Level 4 classification.
    Level 4 'Rejection' Criteria:
    - A solution is already decided (thus not able to apply improvement methodologies).
    - It is a simple just-do-it or straightforward project with solutions like video/print booklets, pamphlets, leaflets, sending reminders, or just reinforcing current processes.
    - It is an IT project (including Excel or programming) that does not involve any actual process change.
    - There is no measurement or documentation of standard work or results.
    - It is a trial or assessment without a concrete implementation.
    - It involves research, randomized control trials, or studies that have pre-empted an intervention.
    - It is just the fine-tuning of new services during a setting-up phase.
    - It's a service development plan like buying new equipment or starting a new service.
    - It's a Regular operation review or fine-tuning (e.g., PAR level, slot allocations, schedules).
    - It's an RCA (Root Cause Analysis) without meaningful supporting data or for a single incident.
    - It's an implementation of Evidence-based practice without measurement.
    """

    prompt = f"""
    You are the Pre-Screening Agent for a hospital's continuous improvement assessment pipeline.
    Evaluate the provided project (JSON data and full PDF) against the Level-4 Exclusionary Rules.

    RULES:
    {level_4_rules}

    EXTRACTED JSON SUMMARY:
    {extracted_data}

    INSTRUCTIONS:
    If you find any violations, you must state the criterion that was met and provide the specific evidence or quote from the document that proves it. If you find no violations, your output must be the single sentence: "No Level 4 criteria violations found."
    """

    # Generate content with both the PDF and the Prompt
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[
            types.Part.from_bytes(data=doc_data, mime_type='application/pdf'),
            prompt
        ],
        config={
            'response_mime_type': 'application/json',
            'response_schema': ScreeningResult,
            'temperature': 0.0  # Zero randomness for auditing
        },
    )

    return response.parsed

# 3. Execution with provided paths
if __name__ == "__main__":
    # Define paths as specified
    json_path = r"C:\Users\liuzh\nuh-qix-assessor\extracted_results\Integrating Artificial Intelligence into Breast Multidiscipl.json"
    pdf_path = r"C:\Users\liuzh\nuh-qix-assessor\project_test\CY25-R2-012_Integrating Artificial Intelligence into Breast Multidisciplinary Tumor Board by Serene Goh Si Ning.pdf"
    
    try:
        print("Starting Pre-Screening Audit...")
        audit_report = run_pre_screening(json_path, pdf_path)
        
        if audit_report.is_eligible:
            print("✅ ELIGIBLE: Project meets baseline requirements for judging.")
        else:
            print(f"❌ INELIGIBLE: Level-4 Violation Detected: {audit_report.primary_violation}")
            print("\nAudit Details:")
            for check in audit_report.detailed_audit:
                status = "🚩 VIOLATION" if check.violation_found else "✅ OK"
                print(f"- {check.criterion}: {status}")
                print(f"  Evidence: {check.evidence_found}")

    except FileNotFoundError as e:
        print(f"File Error: Please ensure paths are correct. {e}")
    except Exception as e:
        print(f"Audit Error: {e}")