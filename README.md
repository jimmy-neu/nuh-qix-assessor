
# NUH-QIX Assessor 🏥

An automated pipeline for clinical project assessment 

## 📌 Overview

The **NUH-QIX Assessor** is designed to streamline the evaluation of clinical project submissions. It replaces manual, error-prone reviews with a deterministic, agentic workflow that ensures every project meets the hospital's "Level 4" rigorous standards before reaching human judges.

---

## 🚀 Installation

Ensure you have Python 3.9+ installed, then run the following commands:

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/jimmy-neu/nuh-qix-assessor.git](https://github.com/jimmy-neu/nuh-qix-assessor.git)
2. **Navigate to the project folder:**
    ```bash
    cd nuh-qix-assessor
3. **Install dependencies:**
```bash
    pip install -r requirements.txt
```



---

## 🤖 Pipeline Architecture

The system operates via multiple specialized agents to bridge the gap between unstructured clinical data and formal auditing logic. (Note: For conversion of pptx file into pdf file use PPTX_PDF.py just need to change the folder path)

### 1. Extraction Agent

The **Extraction Agent** utilizes Vision-Language Models (VLM) to ingest PDFs or slide decks. It maps unstructured text and charts into a strictly defined JSON schema.

* **Why this matters:** Direct LLM processing of complex PDFs often results in "logical drift" or numerical hallucinations.
* **Outcome:** Converts clinical narratives into deterministic facts.

#### Sample JSON Output:

```json
{
  "project_title": "Increasing Clinically Appropriate Contrast-enhanced MRIs (CE MRI) in CKD 4 Patients",
  "department": "Department of Diagnostic Imaging",
  "category": "Process Excellence",
  "problem_statement": "There were limited safe contrast-enhanced imaging options for CKD 4 patients in NUH, leading to the historical avoidance of Contrast-enhanced MRIs (CE MRI) due to concerns about Nephrogenic Systemic Fibrosis (NSF) with older-generation gadolinium-based contrast agents (GBCAs). This occurred despite updated guidelines from the American College of Radiology (ACR) in 2020 and 2024 stating that current-generation GBCAs used in NUH pose minimal risk of NSF for CKD 4 patients, making CE MRI a safer and more appropriate option.",
  "smart_goals": "To increase the proportion of clinically appropriate Contrast-enhanced MRIs (CE MRIs) in CKD 4 patients from a pre-intervention median of 6.7% to 30% within the next 6 months.",
  "methodology": [
    "Gap Analysis",
    "Value Stream Map",
    "Workflow Creation",
    "Education/Training",
    "Standardized Counselling Template",
    "Radiologist Roster Management",
    "System Reminders (EPIC)"
  ],
  "key_results": "The proportion of clinically appropriate Contrast-enhanced MRIs (CE MRIs) in CKD 4 patients increased from a pre-intervention median of 6.7% to a post-intervention median of 24%. Benefits include increased patient satisfaction, reduced patient re-visits, increased clinician diagnostic confidence, reduced resource utilization (hospital beds, scan slots, outpatient clinic slots), and cost savings. Specific cost savings include $70.8 saved per patient per month from reaching diagnosis and reducing follow-up, and $100-$133.3 saved per patient per month by avoiding further investigation with alternative modalities (e.g., Outpatient MRI: $800, Percutaneous biopsy: $400, PET-CT: $1600, Repeat clinic visit: $50).",
  "follow_up_plan": "The plan includes the eventual inclusion of both CKD 4 and CKD 5 patients, with a goal to reconvene with the Nephrology department in 1H 2026 with data. The established workflow can also be introduced to AH and NTFGH diagnostic imaging departments in the future."
}

```

### 2. Pre-screening Agent

The **Pre-screening Agent** acts as an automated gatekeeper. It audits the extracted JSON and the project file together against specific **Level 4 exclusionary criteria** (e.g., filtering out pre-decided interventions or simple IT-only changes).

#### Terminal Output Examples:

**✅ Eligible Project:**

```bash
> python Pre_Screening_agent.py
Starting Pre-Screening Audit...
✅ ELIGIBLE: Project meets baseline requirements for judging.

```

**❌ Ineligible Project (Level-4 Violation):**

```bash
> python Pre_Screening_agent.py
Starting Pre-Screening Audit...
❌ INELIGIBLE: Level-4 Violation Detected: A solution is already decided.

Audit Details:
- A solution is already decided (thus not able to apply improvement methodologies).: 🚩 VIOLATION
  Evidence: The project title is "Integrating Artificial Intelligence into Breast Multidisciplinary Tumor Board". The project's goal, as stated on page 3, is "To enhance the workflow processes of Breast Multi-disciplinary Tumor Board meeting within NUH by leveraging AI tools for data visualization and simplifying workflow processes in 6 months." This explicitly states the solution (leveraging AI tools) as part of the goal, indicating it was decided prior to the application of improvement methodologies to define the solution.
```
### 3. Grading Agent (Multi-Agent Debate)

The final stage employs a **Multi-Agent Debate Architecture** using multimodal AI (processing both the raw PDF and structured JSON simultaneously). 

This ensures that the "Outstanding" (85+) tier remains highly exclusive. By shifting the burden of proof, the system forces the AI to find undeniable, cited evidence for every rubric requirement before awarding high marks.

* **Positive Assessor (Defense):** Actively searches the documents to highlight strengths, locate charts, and argue for maximum points.
* **Negative Assessor (Prosecution):** Acts as a ruthless auditor. It actively attacks vague claims, looks for missing long-term data, and penalizes subjective statements.
* **Independent Judge:** Weighs the debate against the exact hospital rubric. It defaults to lower scores ("Meet Expectations") unless the Defense provides flawless, extracted evidence.

#### Terminal Output

```bash
> python Grading_agent.py
Starting Multi-Agent Grading for: Reducing Arrival to triage wait time for Children's Emergency (R.A.C.E.)
Uploading C:\...\22. Reducing Arrival to triage wait time for Children's Emergency (R.A.C.E.) by Kyi Kyi copy.pdf to Gemini...

PDF Uploaded Successfully!
-> Positive Assessor (Defense) analyzing...
-> Negative Assessor (Prosecution) analyzing...
-> Independent Judge finalizing scores...

✅ Assessment Complete! Final Score: 76/100 (Merit)
Cleaning up PDF from Gemini servers...
Cleanup successful.
```

Sample Graded Output (JSON)
The output provides granular transparency, explaining exactly why a project lost points due to the Skeptic agent's findings.

```json
{
  "assessments": [
    {
      "category": "4. Problem Analysis",
      "max_score": 20,
      "ai_score": 12,
      "ai_justification": "The project applied appropriate lean/PDCA tools (Value Stream Map, P.I.C.K Chart), meeting 'Meet Expectations'. However, the Fishbone diagram lists many potential causes but lacks data-driven prioritization (e.g., Pareto analysis) to confirm which of these are the *main* bottlenecks. As the Skeptic correctly points out, without such evidence, these remain hypotheses rather than 'well identified' root causes, preventing an 'Above Expectations' score.",
      "extracted_quote": "C. Problem Analysis (PLAN) Value Stream Map (PDF Page 4)... Brainstorming P.I.C.K Chart (PDF Page 6)."
    },
    {
      "category": "6. Benefits / Results",
      "max_score": 30,
      "ai_score": 20,
      "ai_justification": "Benefits are quantified (wait times, ePES scores) and sustained for 18 months. However, the rubric explicitly asks: 'Are dips explained?'. The 'ED Arrival to Consult Wait Time' chart (PDF Page 8) shows significant fluctuations and dips in performance post-implementation that are not explained in the text. This lack of explanation undermines the claim of truly *consistent* sustained results required for 'Above Expectations'.",
      "extracted_quote": "Overall, the CE P2 Arrival-to-Consult time has significantly improved and has been sustained for a total of 18 months. (PDF Page 8)"
    }
  ],
  "total_score": 76,
  "label": "Merit"
}
```
