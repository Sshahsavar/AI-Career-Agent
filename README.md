# 💼 Tailored AI Career Agent

An intelligent, multi-tab Streamlit web application powered by **Google Gemini 2.5 Flash** that automatically analyzes job descriptions, optimizes your master resume, selects your top relevant GitHub projects, and drafts structured 4-paragraph cover letters tailored specifically to any job posting.

---

## ✨ Key Features

* **🎯 Match & Gap Analysis:** Mines key requirements from any pasted Job Description or job posting URL, providing an estimated match percentage, key alignment points, and identified skill gaps.
* **📂 Automated Project Selection:** Automatically evaluates your list of GitHub repositories and injects the top 6 most relevant projects directly into your tailored CV.
* **📝 Dynamic Master Resume Builder:** User-friendly UI tabs to manage and edit your Header, Professional Summary, Education, Experience, and Technical Skills without touching raw markdown or JSON.
* **✉️ Structured 4-Paragraph Cover Letters:** Generates highly natural, conversational cover letters strictly following a structured strategy:
  * **Paragraph 1:** Company main objective & key alignment.
  * **Paragraph 2:** Cognitive/mindset trait woven with concrete soft skills examples.
  * **Paragraph 3:** Technical challenge mapping to specific CV tools & projects.
  * **Paragraph 4:** Professional appreciation for the team and culture.
* **💡 Soft Skills Integration:** Dedicated UI tab to manage your soft skills and cognitive abilities, which are dynamically woven into Paragraph 2 of generated cover letters.
* **📄 Compact Exports (PDF & DOCX):** One-click exports for both tailored CVs and cover letters formatted with tight 1-page print margins.
* **⚡ API Token Optimization:** Built-in caching (`@st.cache_data`) and text minification reduce LLM token usage and help prevent 429 rate limit errors on Google's free tier.
* **⚙️ Custom Prompt Engineering:** Tab dedicated to customizing system prompts for CV and Cover Letter generation on the fly.

---

## 🛠️ Project Structure

```text
├── app.py                     # Main Streamlit application
├── requirements.txt           # Python dependencies
├── README.md                  # Project documentation
└── profile_data/              # Local storage directory (auto-created)
    ├── resume_data.json       # Core resume sections
    ├── github_repos.json      # GitHub projects database
    ├── soft_skills.md         # Soft skills & cognitive examples
    └── prompts.json           # System prompt templates# Tailored AI Career Agent & Resume Optimization Engine

An end-to-end, automated AI agent designed to analyze job postings, extract core technical keywords, match user profile assets, and generate tailored, 1-page PDF resumes and cover letters using Google Gemini (`gemini-2.5-flash`) and Streamlit.

---

## Features

* **Job Posting Scraping & Parsing:** Scrapes raw job descriptions directly from live posting URLs or accepts pasted text.
* **Structured Output Mining:** Utilizes Pydantic schemas and Google GenAI SDK to extract keywords, target matches, and skill gaps.
* **Intelligent Project Selection:** Evaluates a local portfolio (`github_repos.json`) and ranks the top 5 most relevant projects for the targeted role.
* **Resume & Cover Letter Tailoring:** Dynamically optimizes professional summary and experience sections while drafting a targeted cover letter that bridges identified experience gaps.
* **Custom 1-Page PDF Generation:** Custom `FPDF` engine with strict markdown formatting and page-overflow validation.
* **Automated Application Tracker:** Logs all applied jobs, company names, URLs, and application dates directly into an Excel workbook (`job_application_tracker.xlsx`).

---

## System Architecture
