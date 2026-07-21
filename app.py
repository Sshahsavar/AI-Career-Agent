import streamlit as st
import json
import os
from datetime import datetime
import pandas as pd
import requests
from bs4 import BeautifulSoup
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
from typing import List, Optional
from dotenv import load_dotenv
from fpdf import FPDF
import re

# Load environment variables from local .env file
load_dotenv()

# --- 1. DEFINING THE AI AGENT'S STRUCTURED OUTPUT ---
class JobAnalysisAgentOutput(BaseModel):
    job_title: str = Field(description="The formal job title extracted from the description.")
    company_name: str = Field(description="The company name extracted from the description.")
    keywords: List[str] = Field(description="Key skills and keywords mined from the JD.")
    matches: List[str] = Field(description="Clear alignment points between the user's profile and the job.")
    gaps: List[str] = Field(description="Missing skills or experience gaps between the user and the job.")
    relevant_repo_names: List[str] = Field(description="Exactly the top 5 most relevant project names chosen from the provided GitHub JSON.")
    new_intro: str = Field(description="Optimized professional summary paragraph. DO NOT include any Markdown headers like '## Summary'.")
    optimized_experience: str = Field(description="Optimized Experience section. You MUST use exactly this format with line breaks:\n**Job Title** | Company | Dates\n* Bullet point 1\n* Bullet point 2\n\nSeparate each job with a blank line. DO NOT include '## Experience' header.")
    tailored_cover_letter: str = Field(description="The core body of the cover letter. Must explicitly bridge the identified gaps.")
    success_probability: str = Field(description="Estimated match level (e.g., '85% Match').")

# --- 2. HYPER-OPTIMIZED 1-PAGE PDF RESUME ENGINE ---
class MarkdownResumePDF(FPDF):
    def __init__(self):
        super().__init__()
        self.set_margins(12, 10, 12)
        self.add_page()
        
    def add_markdown_content(self, markdown_text):
        cleaned_text = markdown_text.replace("{{", "").replace("}}", "")
        cleaned_text = cleaned_text.replace("–", "-").replace("—", "-")
        cleaned_text = cleaned_text.replace("“", '"').replace("”", '"')
        cleaned_text = cleaned_text.replace("‘", "'").replace("’", "'")  
        
        lines = cleaned_text.split("\n")
        
        for line in lines:
            line = line.strip()
            if not line:
                self.ln(0.8) 
                continue
                
            self.set_x(12)
                
            if line.startswith("# "):
                self.set_font("Helvetica", "B", 17)
                self.set_text_color(26, 54, 93) 
                self.cell(0, 7, line[2:], ln=True, align="C")
                
            elif line.startswith("## "):
                self.ln(1.5)
                self.set_font("Helvetica", "B", 11)
                self.set_text_color(43, 108, 176) 
                
                if "(" in line and ")" in line:
                    main_title = line[3:].split("(")[0].strip().upper()
                    note = "(" + line.split("(")[1]
                    
                    title_width = self.get_string_width(main_title + " ")
                    self.cell(title_width, 4.5, main_title + " ", ln=False)
                    
                    self.set_font("Helvetica", "I", 8.5) 
                    self.cell(0, 4.5, note, ln=True)
                else:
                    self.cell(0, 4.5, line[3:].upper(), ln=True)
                    
                self.set_draw_color(226, 232, 240)
                self.line(12, self.get_y(), 198, self.get_y()) 
                self.ln(1)
                
            elif line.startswith("### "):
                tech_match = re.search(r'^(### .*?)\s*\((.*?)\)$', line)
                if tech_match:
                    project_title = tech_match.group(1)[4:] 
                    tech_stack = f"  |  {tech_match.group(2)}" 
                    
                    self.set_font("Helvetica", "B", 10)
                    self.set_text_color(45, 55, 72) 
                    title_width = self.get_string_width(project_title)
                    self.cell(title_width, 4.5, project_title, ln=False)
                    
                    self.set_font("Helvetica", "I", 8.5) 
                    self.set_text_color(74, 85, 104) 
                    self.cell(0, 4.5, tech_stack, ln=True)
                else:
                    self.set_font("Helvetica", "B", 10)
                    self.set_text_color(45, 55, 72) 
                    self.cell(0, 4.5, line[4:], ln=True)

            elif line.startswith("**"):
                self.ln(1.5) 
                clean_line = line.replace("**", "") 
                self.set_font("Helvetica", "B", 10)
                self.set_text_color(45, 55, 72) 
                self.cell(0, 4.5, clean_line, ln=True)
                
            elif line.startswith("*") or line.startswith("•") or line.startswith("-"):
                self.set_font("Helvetica", "", 9) 
                self.set_text_color(45, 55, 72)
                self.set_x(17) 
                bullet_content = re.sub(r'^[*•-]\s*', '', line)
                self.cell(3, 4.0, chr(149)) 
                self.multi_cell(0, 4.0, bullet_content) 
                
            else:
                self.set_text_color(45, 55, 72)
                if "shahab.shvr@gmail.com" in line or "LinkedIn:" in line or "647-410-5827" in line:
                    self.set_font("Helvetica", "", 9)
                    self.set_text_color(74, 85, 104) 
                    self.cell(0, 4.0, line, ln=True, align="C")
                else:
                    self.set_font("Helvetica", "", 9)
                    self.multi_cell(0, 4.0, line)

# --- 3. PURE PYTHON NATIVE COVER LETTER PDF GENERATOR ---
class StructuredCoverLetterPDF(FPDF):
    def __init__(self):
        super().__init__()
        self.set_margins(20, 20, 20)
        self.add_page()
        
    def generate_layout(self, body_content, company_name):
        self.set_font("Helvetica", "", 11)
        self.set_text_color(45, 55, 72)
        
        # --- NEW: Safe Unicode text scrubbing for Cover Letter ---
        body_content = body_content.replace("–", "-").replace("—", "-")
        body_content = body_content.replace("“", '"').replace("”", '"')
        body_content = body_content.replace("‘", "'").replace("’", "'")
        
        if company_name:
            company_name = company_name.replace("–", "-").replace("—", "-")
            company_name = company_name.replace("“", '"').replace("”", '"')
            company_name = company_name.replace("‘", "'").replace("’", "'")
            
        salutation = f"Dear Hiring Team at {company_name}," if company_name else "Dear Hiring Manager,"
        self.cell(0, 6, salutation, ln=True)
        self.ln(5)
        
        self.multi_cell(0, 6, body_content)
        self.ln(6)
        
        self.cell(0, 6, "Thanks for your time,", ln=True)
        self.ln(4)
        
        self.set_font("Helvetica", "B", 11)
        self.cell(0, 5, "Shahab Shahsavar", ln=True)
        self.set_font("Helvetica", "", 10)
        self.set_text_color(74, 85, 104)
        self.cell(0, 5, "Toronto, ON | 647-410-5827 | shahab.shvr@gmail.com", ln=True)
        self.ln(6)
        
        self.set_text_color(45, 55, 72)
        today_str = datetime.today().strftime('%B %d, %Y')
        self.cell(0, 5, today_str, ln=True)

# --- 4. SCRAPING AND TRACKING LOGIC ---
def extract_text_from_url(url):
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)'}
        response = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')
        for script in soup(["script", "style"]):
            script.decompose()
        return soup.get_text(separator=' ', strip=True)
    except Exception as e:
        st.error(f"Web scraper connection timeout or failure: {e}")
        return None

def update_excel_tracker(job_title, company, url):
    excel_file = "job_application_tracker.xlsx"
    today = datetime.today().strftime('%Y-%m-%d')
    new_data = pd.DataFrame([{
        "Date Applied": today,
        "Company": company,
        "Job Title": job_title,
        "URL": url if url else "Pasted Text",
        "Status": "Tailored / Ready to Apply"
    }])
    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file)
        df = pd.concat([df, new_data], ignore_index=True)
    else:
        df = new_data
    df.to_excel(excel_file, index=False)

# --- 5. STREAMLIT FRAMEWORK UI ---
st.set_page_config(page_title="Tailored AI Career Agent", layout="wide")
st.title("💼 Tailored AI Career Agent")

st.sidebar.header("📁 System Status Check")
if os.path.exists("profile_data/resume_master.md") and os.path.exists("profile_data/github_repos.json"):
    st.sidebar.success("Master profile configurations loaded!")
else:
    st.sidebar.error("Error loading master profile data configurations from 'profile_data/' directory.")

if "agent_results" not in st.session_state:
    st.session_state.agent_results = None

input_type = st.radio("Choose source profile input pipeline:", ["Paste Job Description Text", "Provide Job Posting URL"])
job_description = ""
job_url = ""

if input_type == "Paste Job Description Text":
    job_description = st.text_area("Paste Job Description Here:", height=250)
else:
    job_url = st.text_input("Enter Active Job Posting Link URL:")
    if job_url:
        with st.spinner("Extracting text metrics from pipeline target..."):
            job_description = extract_text_from_url(job_url)

if st.button("🚀 Execute Agent Target Matching") and job_description:
    with st.spinner("Executing 5-Step Profile Optimization..."):
        
        with open("profile_data/resume_master.md", "r") as f:
            master_resume = f.read()
        with open("profile_data/github_repos.json", "r") as f:
            github_repos = json.load(f)
            
        client = genai.Client()
        
        prompt = f"""
        You are an expert career agent. Execute the following steps strictly based on the provided Job Description, Master Resume, and GitHub Repositories.

        1. MINE KEYWORDS: Extract core skills and requirements from the Job Description.
        2. SORT PROJECTS: Analyze the USER GITHUB REPOSITORIES. Select exactly the top 5 most relevant projects for this specific role.
        3. OPTIMIZE RESUME: Optimize both the 'Professional Summary' and the 'Experience' sections of the Master Resume to highlight relevant skills.
        4. GENERATE COVER LETTER: Draft the body of a cover letter based on the optimized resume. Explicitly use the cover letter to bridge and address any gaps between the resume and the job description.
        5. MATCH & GAP ANALYSIS: Determine the success probability, list clear matches, and list specific gaps.

        JOB DESCRIPTION:
        {job_description}
        
        USER MASTER RESUME:
        {master_resume}
        
        USER GITHUB REPOSITORIES:
        {json.dumps(github_repos)}
        """
        
        try:
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=JobAnalysisAgentOutput,
                    temperature=0.2
                ),
            )
        except Exception as e:
            st.error("⚠️ Google's servers are currently experiencing a high volume of traffic. Please wait 5 seconds and click the button again to retry.")
            st.stop()
        
        result = json.loads(response.text)
        update_excel_tracker(result['job_title'], result['company_name'], job_url)
        
        project_markdown_blocks = []
        for repo_name in result['relevant_repo_names'][:5]:  
            match = next((r for r in github_repos if r['name'].lower() == repo_name.lower()), None)
            if match:
                bullets_str = "\n".join([f"* {b}" for b in match['bullets']])
                tech = ", ".join(match['tech_stack'])
                block = f"### {match['name']} ({tech})\n{match['description']}\n{bullets_str}\n"
                project_markdown_blocks.append(block)
        
        final_projects_section = "\n".join(project_markdown_blocks)
        
        clean_intro = re.sub(r'^#+.*?\n', '', result['new_intro'] + '\n', flags=re.IGNORECASE).strip()
        clean_exp = re.sub(r'^#+.*?\n', '', result['optimized_experience'] + '\n', flags=re.IGNORECASE).strip()
        
        final_resume_content = master_resume
        final_resume_content = re.sub(r'##\s*(?:Professional\s+)?Summary\s*\{\{.*?\}\}', f"## PROFESSIONAL SUMMARY\n{clean_intro}", final_resume_content, flags=re.DOTALL | re.IGNORECASE)
        final_resume_content = re.sub(r'##\s*(?:Quantitative\s+)?Projects\s*\{\{.*?\}\}', f"## QUANTITATIVE PROJECTS (Please check my Github for more)\n{final_projects_section}", final_resume_content, flags=re.DOTALL | re.IGNORECASE)
        final_resume_content = re.sub(r'##\s*(?:Professional\s+)?Experience\s*\{\{.*?\}\}', f"## PROFESSIONAL EXPERIENCE\n{clean_exp}", final_resume_content, flags=re.DOTALL | re.IGNORECASE)
        
        final_resume_content = final_resume_content.replace("{{", "").replace("}}", "")
        
        os.makedirs("output", exist_ok=True)
        safe_company = result['company_name'].replace(' ', '_').replace('/', '_')
        safe_title = result['job_title'].replace(' ', '_').replace('/', '_')
        
        pdf_path = f"output/{safe_company}_{safe_title}_Resume.pdf"
        cl_path = f"output/{safe_company}_{safe_title}_Cover_Letter.pdf"
        
        pdf_generator = MarkdownResumePDF()
        pdf_generator.add_markdown_content(final_resume_content)
        pdf_generator.output(pdf_path)
        resume_pages = pdf_generator.page_no()
        
        cl_generator = StructuredCoverLetterPDF()
        cl_generator.generate_layout(result['tailored_cover_letter'], result['company_name'])
        cl_generator.output(cl_path)
        cl_pages = cl_generator.page_no()
        
        st.session_state.agent_results = {
            "pdf_path": pdf_path,
            "cl_path": cl_path,
            "safe_company": safe_company,
            "success_probability": result['success_probability'],
            "keywords": result['keywords'],
            "matches": result['matches'],
            "gaps": result['gaps'],
            "final_resume_content": final_resume_content,
            "resume_pages": resume_pages,
            "cl_pages": cl_pages
        }

if st.session_state.agent_results is not None:
    res = st.session_state.agent_results
    
    st.success(f"🎉 Tailored materials compiled and saved into local storage folder 'output/'!")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="Estimated Alignment Score", value=res['success_probability'])
        st.write("**Identified Keywords:**", ", ".join(res['keywords']))
    with col2:
        st.write("✅ **Profile Matches:**")
        for match in res['matches']:
            st.write(f"- {match}")
    with col3:
        st.write("❌ **Identified Gaps (Addressed in Cover Letter):**")
        if res['gaps']:
            for gap in res['gaps']:
                st.write(f"- {gap}")
        else:
            st.write("No major gaps identified.")
            
    st.divider()
    
    # --- PAGE COUNT VALIDATION ---
    st.subheader("📄 Document Page Verification")
    page_col1, page_col2 = st.columns(2)
    with page_col1:
        if res['resume_pages'] == 1:
            st.success(f"✅ Resume Length: 1 Page")
        else:
            st.error(f"⚠️ Resume Length: {res['resume_pages']} Pages. Text is overflowing.")
            
    with page_col2:
        if res['cl_pages'] == 1:
            st.success(f"✅ Cover Letter Length: 1 Page")
        else:
            st.error(f"⚠️ Cover Letter Length: {res['cl_pages']} Pages. Text is overflowing.")
            
    st.divider()
    
    st.subheader("📥 Action Dashboard Download Panel")
    
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        with open(res['pdf_path'], "rb") as pdf_file:
            st.download_button("📄 Download Custom Resume (PDF Format)", pdf_file, file_name=f"{res['safe_company']}_Resume.pdf", mime="application/pdf")
            
    with col_dl2:
        with open(res['cl_path'], "rb") as cl_file:
            st.download_button("✉️ Download Target Cover Letter (PDF Format)", cl_file, file_name=f"{res['safe_company']}_Cover_Letter.pdf", mime="application/pdf")
    
    with st.expander("Show Generated File Layout Context"):
        st.code(res['final_resume_content'], language="markdown")
        
    st.divider()
    
    st.write("🔄 **Finish Session Configuration**")
    if st.button("Reset Dashboard & Clear Session"):
        st.session_state.agent_results = None
        st.rerun()
