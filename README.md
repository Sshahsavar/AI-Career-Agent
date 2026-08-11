# Tailored AI Career Agent & Resume Optimization Engine

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
