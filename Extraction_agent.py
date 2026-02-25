import os
import httpx
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
from dotenv import load_dotenv
import re

# 1. Load the environment variables from the hidden.env file
load_dotenv()

# 2. Initialize the client (it will securely pull the key from your system)
client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))

class ProjectExtraction(BaseModel):
    project_title: str = Field(description="The title of the quality improvement project.")
    department: str = Field(description="The hospital department submitting the project.")
    category: str = Field(description="The framework category: 6S, Process Excellence, or Service Experience.")
    problem_statement: str = Field(description="The core problem or initial state being addressed.")
    smart_goals: str = Field(description="Specific, measurable, achievable, realistic, and time-bound goals.")
    methodology: list[str] = Field(description="List of Lean, Design Thinking, or PDCA tools applied (e.g., Value Stream Map, Gap Analysis).")
    key_results: str = Field(description="Quantitative and qualitative benefits, financial savings, or sustained improvements extracted from text or charts.")
    follow_up_plan: str = Field(description="Plans to sustain the results or spread the implementation to other departments.")

def extract_clinical_project(local_pdf_path: str) -> ProjectExtraction:
    """
    Reads a local clinical project submission (PDF) and extracts structured data.
    """
    # Read the local file in binary mode
    with open(local_pdf_path, "rb") as file:
        doc_data = file.read()

    prompt = """
    You are the Extraction Agent for a hospital's continuous improvement assessment pipeline.
    Analyze the provided project submission document. Your task is to extract the relevant information 
    to populate the required schema. Pay close attention to both unstructured narrative text and 
    visual elements like charts, graphs, and tables to capture the full scope of the methodologies and results.
    """

    # Generate content using Gemini Flash with Structured Outputs
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[
            types.Part.from_bytes(data=doc_data, mime_type='application/pdf'),
            prompt
        ],
        config={
            'response_mime_type': 'application/json',
            'response_schema': ProjectExtraction,
            'temperature': 0.1 
        },
    )

    return response.parsed

if __name__ == "__main__":
    input_folder = "./project_test"
    output_folder = "./extracted_results"
    
    # 1. Create the new folder if it doesn't already exist
    os.makedirs(output_folder, exist_ok=True)
    
    for filename in os.listdir(input_folder):
        if filename.lower().endswith(".pdf"):
            local_pdf_path = os.path.join(input_folder, filename)
            
            try:
                # Extract the data
                extracted_data = extract_clinical_project(local_pdf_path)
                
                # 2. Create a proper, safe filename from the extracted project title
                # Remove characters that are invalid in Windows/Mac filenames
                safe_title = re.sub(r'[\\/*?:"<>|]', "", extracted_data.project_title)
                
                # Limit the filename length just in case the title is extremely long
                safe_filename = f"{safe_title[:60].strip()}.json"
                output_path = os.path.join(output_folder, safe_filename)
                
                # 3. Save the results into the new folder
                with open(output_path, "w", encoding="utf-8") as json_file:
                    json_file.write(extracted_data.model_dump_json(indent=2))
                    
                print(f"Successfully saved data from {filename} -> {safe_filename}")
                
            except Exception as e:
                print(f"Error processing {filename}: {e}")