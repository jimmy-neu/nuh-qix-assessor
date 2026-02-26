import os
import json
import time
from google import genai
from google.genai import types
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Initialize the new Gemini Client
client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))


MODEL_NAME = "gemini-2.5-flash"

# The FULL NUH Process Improvement Rubric
FULL_RUBRIC = """
Assessment Score (Minimum): Outstanding-85; Merit-70; Recognition - 50

1. Header (Max Score: 5)
- Meet Expectations (3): Are appropriate team members, leader & sponsor identified to participate?
- Above Expectations (5): Does the title clearly explain the goal of the project?

2. Background (Max Score: 10)
- Meet Expectations (6): Is the nature of the problem clearly explained? Is the problem worthwhile resolving? Is the scope clearly defined and manageable for the Improvement Sprints /project length?
- Above Expectations (10): Is appropriate data collected and presented to quantify and understand the problem?

3. Goal (Max Score: 5)
- Meet Expectations (3): Is the goal SMART? (Specific, measurable, achievable, realistic & time based)
- Above Expectations (5): Is the goal linked to NUH / department objectives and to the problem defined?

4. Problem Analysis (Max Score: 20)
- Meet Expectations (12): Are the appropriate lean / PDCA tools applied to analyze the problems? Does the analysis link back to the problem and goals?
- Above Expectations (20): Are the root causes and main delays/bottlenecks of the problem well identified?

5. Implementation Plan (Max Score: 20)
- Meet Expectations (12): Do the solutions and implementation plan directly address the root causes of problems? Is the implementation plan clear and timely? Does it include specific what, who & when?
- Above Expectations (20): For Improvement Sprints, was a rapid experiment carried out and were insights gained from it? Were some changes started immediately during the event or the next week?

6. Benefits / Results (Max Score: 30)
- Meet Expectations (20): Is there a clear significant improvement? Do the results match the goal? Are the benefits quantified? Are there intangible results? Are clear charts used to show results/improvement? Are dips explained?
- Above Expectations (30): Are the results sustained over the time? (Ideally show results >= 3 months)

7. Follow-up, Spread & Insights (Max Score: 10)
- Meet Expectations (6): Has the team created a plan to overcome any outstanding or potential issues? Has the team created and/or implemented an effective plan to spread the project?
- Above Expectations (10): Has reflection taken place and insights / lessons learned been identified?
"""

def upload_pdf_to_gemini(pdf_path):
    """Uploads the PDF using the new Client.files API."""
    print(f"Uploading {pdf_path} to Gemini...")
    pdf_file = client.files.upload(
        file=pdf_path, 
        config={'mime_type': 'application/pdf'}
    )
    
    # Wait for the file to finish processing in the Gemini backend
    while pdf_file.state.name == "PROCESSING":
        print(".", end="", flush=True)
        time.sleep(2)
        pdf_file = client.files.get(name=pdf_file.name)
    
    if pdf_file.state.name == "FAILED":
        raise ValueError(f"Failed to process PDF: {pdf_file.name}")
        
    print("\nPDF Uploaded Successfully!")
    return pdf_file

def call_gemini_agent(system_instruction, pdf_file, text_prompt, require_json=False):
    """Helper function to call Gemini using the new models.generate_content API."""
    
    # Configure the generation parameters using types
    config = types.GenerateContentConfig(
        system_instruction=system_instruction,
        temperature=0.2,
        response_mime_type="application/json" if require_json else "text/plain"
    )
    
    # Pass both the uploaded PDF file object and the text prompt
    contents = [pdf_file, text_prompt]
    
    response = client.models.generate_content(
        model=MODEL_NAME,
        contents=contents,
        config=config
    )
    return response.text

def positive_assessor(pdf_file, json_text):
    print("-> Positive Assessor (Defense) analyzing...")
    prompt = f"""You are the Positive Advocate for this clinical project. 
    Using BOTH the provided PDF document and the JSON summary, review the project against this rubric:
    {FULL_RUBRIC}
    
    Your goal is to highlight all strengths. Find specific evidence, charts, or quotes in the PDF and JSON that justify 'Above Expectations' scores for every category. Formulate a strong defensive argument."""
    
    # We prefix the prompt with a clear label for the JSON data
    text_prompt = f"EXTRACTED PROJECT JSON:\n{json_text}\n\n{prompt}"
    return call_gemini_agent("Act as a strict but highly supportive positive advocate.", pdf_file, text_prompt)
def negative_assessor(pdf_file, json_text):
    print("-> Negative Assessor (Prosecution) analyzing...")
    prompt = f"""You are a ruthless, veteran clinical auditor known for extreme strictness. 
    Using BOTH the provided PDF document and the JSON summary, review the project against this rubric:
    {FULL_RUBRIC}
    
    Your goal is to aggressively drag the score down. You must actively search for technicalities, missing long-term data, or subjective claims. 
    - If a criterion requires multiple elements (e.g., 'quantified benefits' AND 'intangible results'), and they only have one, aggressively attack the missing element.
    - If they claim 'sustained results', strictly verify if the charts in the PDF explicitly prove >= 3 months. If it's only 2.5 months, flag it as a failure.
    - Argue fiercely why this project DOES NOT deserve 'Above Expectations' and must be capped at 'Meet Expectations' or 'Below Expectations'."""
    
    text_prompt = f"EXTRACTED PROJECT JSON:\n{json_text}\n\n{prompt}"
    return call_gemini_agent("Act as a hostile, highly critical project auditor who penalizes missing data heavily.", pdf_file, text_prompt)

def independent_judge(pdf_file, json_text, pos_arg, neg_arg):
    print("-> Independent Judge finalizing scores...")
    system_prompt = f"""You are the Lead Meta-Judge for the NUH QIX awards. 
    You have the original PDF submission, the extracted JSON, a Positive Advocate's review, and a Strict Skeptic's review.
    
    RUBRIC & THRESHOLDS:
    {FULL_RUBRIC}
    - Outstanding: >= 85 (HARD REQUIREMENT: Only the top 10% of elite projects achieve this. Evidence must be flawless.)
    - Merit: >= 70
    - Recognition: >= 50
    
    CRITICAL GRADING RULES:
    1. THE BURDEN OF PROOF: You MUST default to 'Meet Expectations' or lower. You are strictly forbidden from awarding 'Above Expectations' unless the Positive Advocate provides explicit, undeniable quotes/charts from the source documents that satisfy EVERY SINGLE requirement in that rubric tier.
    2. PENALIZE VAGUENESS: If the Skeptic successfully points out that a claim is subjective, unquantified, or lacks a specific timeframe, you MUST downgrade the score. 
    3. EXACT DISCRETE SCORES: You must only select the exact integer scores provided in the rubric (e.g., for Background, you must choose 3, 6, or 10. Do not invent a score like 8 or 9).
    
    INSTRUCTIONS:
    Cross-reference the debate. Rule on each category. Output strictly matching this JSON schema:
    {{
      "assessments": [
        {{
          "category": "String (e.g., '1. Background')",
          "max_score": "Integer (e.g., 10)",
          "ai_score": "Integer (Must be an exact discrete score from the rubric)",
          "ai_justification": "String (Explain why you rejected the higher score, or why the evidence was so flawless it forced you to award it)",
          "extracted_quote": "String (Exact quote from the PDF/JSON supporting the score)"
        }}
      ]
    }}"""
    
    debate_context = f"""
    --- EXTRACTED JSON ---
    {json_text}
    
    --- POSITIVE ADVOCATE ARGUMENT ---
    {pos_arg}
    
    --- STRICT SKEPTIC ARGUMENT ---
    {neg_arg}
    """
    
    raw_json = call_gemini_agent(system_prompt, pdf_file, debate_context, require_json=True)
    return json.loads(raw_json)

def grade_project(pdf_filepath, json_filepath):
    # 1. Load JSON
    with open(json_filepath, 'r') as f:
        project_json = json.load(f)
        json_text = json.dumps(project_json, indent=2)
        
    print(f"Starting Multi-Agent Grading for: {project_json.get('project_title', 'Unknown')}")
    
    # 2. Upload PDF to Gemini
    pdf_file = upload_pdf_to_gemini(pdf_filepath)
    
    try:
        # 3. Run Debate
        pos_arg = positive_assessor(pdf_file, json_text)
        neg_arg = negative_assessor(pdf_file, json_text)
        
        # 4. Final Judgment
        final_assessment = independent_judge(pdf_file, json_text, pos_arg, neg_arg)
        
        # 5. Calculate Score
        total_score = sum(int(item['ai_score']) for item in final_assessment['assessments'])
        
        if total_score >= 85:
            label = "Outstanding"
        elif total_score >= 70:
            label = "Merit"
        elif total_score >= 50:
            label = "Recognition"
        else:
            label = "Below Recognition"
            
        final_assessment["total_score"] = total_score
        final_assessment["label"] = label
        
        print(f"\n✅ Assessment Complete! Final Score: {total_score}/100 ({label})")
        return final_assessment
        
    finally:
        # 6. Clean up the file from Gemini's servers after processing
        print("Cleaning up PDF from Gemini servers...")
        client.files.delete(name=pdf_file.name)
        print("Cleanup successful.")

if __name__ == "__main__":
    # Update these paths to point to your actual local files
    test_pdf = r"C:\Users\liuzh\nuh-qix-assessor\project_test\22. Reducing Arrival to triage wait time for Children's Emergency (R.A.C.E.) by Kyi Kyi copy.pdf"
    test_json = r"C:\Users\liuzh\nuh-qix-assessor\extracted_results\Reducing Arrival to triage wait time for Children's Emergenc.json" 
    
    if os.path.exists(test_pdf) and os.path.exists(test_json):
        result = grade_project(test_pdf, test_json)
        
        with open("graded_result.json", "w") as out:
            json.dump(result, out, indent=4)
    else:
        print("Error: Could not find the specified PDF or JSON file. Please check the paths.")