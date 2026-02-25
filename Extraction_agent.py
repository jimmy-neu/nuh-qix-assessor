import os
import httpx
from google import genai
from google.genai import types
from pydantic import BaseModel, Field

# 1. Initialize the client
# Ensure your GEMINI_API_KEY environment variable is set
client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))

# 2. Define the exact schema the Extraction Agent must return using Pydantic
class ProjectExtraction(BaseModel):
    project_title: str = Field(description="The title of the quality improvement project.")
    department: str = Field(description="The hospital department submitting the project.")
    category: str = Field(description="The framework category: 6S, Process Excellence, or Service Experience.")
    problem_statement: str = Field(description="The core problem or initial state being addressed.")
    smart_goals: str = Field(description="Specific, measurable, achievable, realistic, and time-bound goals.")
    methodology: list[str] = Field(description="List of Lean, Design Thinking, or PDCA tools applied (e.g., Value Stream Map, Gap Analysis).")
    key_results: str = Field(description="Quantitative and qualitative benefits, financial savings, or sustained improvements extracted from text or charts.")
    follow_up_plan: str = Field(description="Plans to sustain the results or spread the implementation to other departments.")

def extract_clinical_project(pdf_url: str) -> ProjectExtraction:
    """
    Downloads a clinical project submission (PDF) and extracts structured data.
    """
    # Fetch the raw document bytes
    doc_data = httpx.get(pdf_url).content

    prompt = """
    You are the Extraction Agent for a hospital's continuous improvement assessment pipeline.
    Analyze the provided project submission document. Your task is to extract the relevant information 
    to populate the required schema. Pay close attention to both unstructured narrative text and 
    visual elements like charts, graphs, and tables to capture the full scope of the methodologies and results.
    """

    # 3. Generate content using Gemini 2.0 Flash with Structured Outputs
    response = client.models.generate_content(
        model='gemini-2.0-flash',
        contents=[
            types.Part.from_bytes(data=doc_data, mime_type='application/pdf'),
            prompt
        ],
        config={
            'response_mime_type': 'application/json',
            'response_schema': ProjectExtraction,
            'temperature': 0.1 # Keep temperature low for factual extraction
        },
    )

    # 4. Return the parsed, validated Pydantic object
    return response.parsed

# --- Example Usage ---
if __name__ == "__main__":
    # Example URL pointing to a hospital project submission PDF
    sample_pdf_url = "https://example-hospital.com/submissions/QIX_Diagnostic_Imaging.pdf"
    
    try:
        extracted_data = extract_clinical_project(sample_pdf_url)
        print(f"Project Name: {extracted_data.project_title}")
        print(f"Department: {extracted_data.department}")
        print(f"Results Found: {extracted_data.key_results}")
    except Exception as e:
        print(f"Extraction failed: {e}")