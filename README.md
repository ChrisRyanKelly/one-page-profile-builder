# One Page Profile Builder (Apps Script × Gemini)

**An intelligent automation tool designed to streamline One Page Profile creation, reduce duplication, and improve consistency across staff.**

---

## Problem

The existing process for creating One Page Profiles was inefficient, inconsistent, and wasteful. Staff would typically interview a student and take handwritten notes, then retype and format those notes into a Word template before saving and uploading the final document to **Microsoft Teams**. This duplication of effort led to unnecessary time pressures and frequent formatting inconsistencies between staff members. Additionally, the reliance on paper notes created avoidable waste, as most drafts were printed and later discarded. The overall workflow was functional but far from streamlined, leaving significant room for improvement.

---

## Solution

This project introduces an automated and standardised workflow for creating One Page Profiles using **Google Forms**, **Google Sheets**, and **Apps Script**. The new system removes the need for handwritten notes and manual transcription by allowing staff to record student responses directly into a digital form. Once submitted, the information is processed automatically: an integrated **AI assistant** refines the raw text into clear, professional UK English and structures it within a pre-formatted template. The resulting **Google Doc** provides a consistent, high-quality draft that staff can quickly review, edit, and export.

This approach streamlines the process, saves staff time, and ensures every profile is produced with clarity, consistency, and professionalism.

---

## Key Features

- Automated AI-powered text generation using **Gemini**
- Customisable **Google Docs** template for consistent formatting
- Secure data handling with automatic name redaction
- Built-in fallback for AI downtime to ensure reliability
- One-click export to **PDF** or **Word** for upload to **Microsoft Teams**
- Fully digital workflow that eliminates paper waste

---

## Workflow

1. Enter student responses directly into a **Google Form**.
    
   <img src="https://github.com/ChrisRyanKelly/one-page-profile-builder/blob/master/public/assets/form.png" width="600" alt="Goolge-Form" />
    
2. Responses are captured automatically in a **Google Sheet**.
    
   <img src="https://github.com/ChrisRyanKelly/one-page-profile-builder/blob/master/public/assets/form-responses.png" width="600" alt="Form-Responses" />
    
3. An **Apps Script trigger** processes each submission:
    - Calls **Gemini** to polish responses into clear, professional UK English
    - Generates a draft **Google Doc** from a standardised template
    - Saves a ready-to-edit draft in **Google Drive** for staff QA

   <img src="https://github.com/ChrisRyanKelly/one-page-profile-builder/blob/master/public/assets/drive.png" width="600" alt="Google-Drive" />
        
4. Export the reviewed document as **PDF** or **Word** and upload to **Microsoft Teams**.
    
   <img src="https://github.com/ChrisRyanKelly/one-page-profile-builder/blob/master/public/assets/doc-export.png" width="600" alt="Google-Doc-Export" />
    

---

## Guardrails

The system has been designed with clear guardrails to ensure privacy, reliability, and professional accountability throughout the process. Student data is handled responsibly, with names automatically masked before any text is sent to the AI model to protect confidentiality. Built-in fallback mechanisms ensure that a usable draft is always produced, even if the AI service is temporarily unavailable, maintaining workflow continuity for staff. Most importantly, the process preserves human oversight at every stage. Staff are required to review, edit, and approve each draft before export, ensuring that AI-generated content remains accurate, appropriate, and aligned with the student’s authentic voice.

---

## How It Works

### Configurable Settings (in Apps Script)

```jsx
const TEMPLATE_DOC_ID = '...';       // Google Doc template
const OUTPUT_DOC_FOLDER_ID = '...';  // Folder for generated profiles
const GEMINI_API_PROP_KEY = 'GEMINI_API_KEY'; // API key stored in script properties
```

### Setup

1. **Google Workspace**
    - Create a **Google Form** for staff to complete.
    - Link it to a **Google Sheet**.
2. **Apps Script**
    - Add the provided script to the **Google Sheet**.
    - Set `TEMPLATE_DOC_ID` and `OUTPUT_DOC_FOLDER_ID`.
    - Store your **Gemini** API key in **Script Properties** (`GEMINI_API_KEY`).
3. **Triggers**
    - Run `createOnFormSubmitTrigger()` once to install the `onFormSubmit` trigger.
4. **Template**
    - Create a **Google Doc** template with placeholders, e.g. `{{StudentName}}`, `{{AboutMe}}`, `{{HowToSupportMe}}`.

---

## Intended Impact

This project is designed to directly address the inefficiencies of the previous workflow by improving speed, consistency, and sustainability in the creation of One Page Profiles. Through automation and AI-assisted drafting, staff can focus on refining content rather than retyping notes. The system standardises formatting and tone across all profiles, ensures each one remains student-centred, and significantly reduces paper usage by enabling a fully digital process.

Overall, the tool supports a more efficient, consistent, and environmentally responsible approach to profile creation while maintaining professional human oversight at every stage.
