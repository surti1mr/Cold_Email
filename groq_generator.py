"""Generate personalized cold emails using Groq (LLaMA 3.3 70B).

Strategy: the email body is assembled in Python from a hardcoded template.
Groq only extracts three values from the job description:
  - job_title
  - company_name
  - subject  (short subject line)
This guarantees the template is followed 100% of the time.
"""

from __future__ import annotations

import json
import os
import re
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent / ".env")

MODEL = "llama-3.3-70b-versatile"
PORTFOLIO_URL = "https://portfolio-ten-pi-50.vercel.app/"

# ── Hardcoded email templates ─────────────────────────────────────────────────
# {job_title} and {company_name} are the only placeholders filled by LLM output.

_TEMPLATE_CMU_ALUMNI = """\
Hi {{name}},

I came across your profile and noticed you're a fellow Central Michigan University alum. Fire Up Chips!

I'm reaching out because I noticed an opening for {job_title} at {company_name} and I'm very interested in the role. I recently completed my M.S. in Information Systems (SAP) at CMU in December 2025 with a 3.95 GPA, and I have 5+ years of software engineering experience with a growing specialization in AI/ML systems including RAG pipelines, LLM integration, FAISS vector search, and full-stack development with React, Next.js, Node.js, and Python (FastAPI).

Most recently, I built FinanceAI, an end-to-end AI-powered personal finance application, as a personal project to deepen my AI engineering skills. I'm currently interning at Detroit Manufacturing Systems in Detroit while actively pursuing full-time opportunities.

On a lighter note, during my time at CMU, I competed in a stand-up comedy competition in Downtown Mount Pleasant and secured 2nd place based on audience voting. I mention it because I think it reflects how I approach problems: with creativity, confidence, and a willingness to think outside the box. Those same qualities carry into how I communicate with teams and tackle technical challenges.

I've attached my resume for your reference, and you can view my full portfolio and projects here: """ + PORTFOLIO_URL + """
Given our shared CMU connection, I'd love to hear your thoughts on the team and the role. Even a quick 15-minute conversation would be incredibly valuable.

Thank you so much for your time, {{name}}!

Best,
Mayank Surti
(248) 704-9118 | linkedin.com/in/mayank-surti-593bb3185/
github.com/surti1mr
Portfolio: """ + PORTFOLIO_URL

_TEMPLATE_RECRUITER = """\
Hi {{name}},

I recently applied for the {job_title} position at {company_name} and wanted to follow up directly. I'm genuinely excited about this opportunity and believe my background is a strong match for what your team is looking for.

A quick snapshot of what I bring:

\u2022 5+ years of software engineering experience in full-stack development (React, Next.js, Node.js, TypeScript) and AI/ML systems (RAG pipelines, LLM integration, FAISS vector search, Python/FastAPI)
\u2022 M.S. in Information Systems from Central Michigan University (3.95 GPA, December 2025)
\u2022 Currently interning at Detroit Manufacturing Systems in Detroit, building production-grade tools and data pipelines
\u2022 Built FinanceAI end-to-end, an AI-powered personal finance app using LLaMA 3.3 70B, FAISS, and Next.js 14, live on Vercel

You can view my full portfolio, including live project demos, here: """ + PORTFOLIO_URL + """

On the softer side, I competed in a stand-up comedy competition during grad school and placed 2nd based on audience voting. I mention it because I think it reflects my ability to communicate clearly, think creatively, and connect with people, skills I bring to every team I work with.

I've attached my resume for your review. I'd love the chance to speak with you and learn more about the role and your team's needs. I'm flexible on timing and happy to connect whenever works best for you.

Thank you for your time. I look forward to hearing from you!

Best regards,
Mayank Surti
(248) 704-9118 | surti1mr@cmich.edu
linkedin.com/in/mayank-surti-593bb3185/ | github.com/surti1mr
Portfolio: """ + PORTFOLIO_URL

_TEMPLATE_IT_EMPLOYEE = """\
Hi {{name}},

I hope this message finds you well! I came across your profile on LinkedIn and noticed you're working at {company_name}. I'm currently applying for the {job_title} role there and wanted to reach out directly.

I'm Mayank Surti, a software engineer with 5+ years of experience specializing in full-stack development (React, Next.js, Node.js, TypeScript) and AI/ML systems including RAG pipelines, LLM integration, FAISS vector search, and Python (FastAPI). I recently completed my M.S. in Information Systems at Central Michigan University (3.95 GPA) and am currently interning at Detroit Manufacturing Systems in Detroit while pursuing full-time opportunities.

Most recently, I built FinanceAI, an end-to-end AI-powered personal finance application, as a personal project to sharpen my AI engineering skills, something I'm genuinely passionate about beyond just the day job. You can see this and other projects on my portfolio: """ + PORTFOLIO_URL + """

On a fun note, I also competed in a stand-up comedy competition during grad school and secured 2nd place based on audience voting. I bring that same energy of creative thinking and confident communication to every team I work with.

I've attached my resume for your reference. I'd love to hear your perspective on the team culture and what it's like working at {company_name}. Even a quick 10-15 minute conversation would be incredibly helpful as I go through the process.

Thank you so much for your time. I really appreciate it!

Best,
Mayank Surti
(248) 704-9118 | linkedin.com/in/mayank-surti-593bb3185/
github.com/surti1mr
Portfolio: """ + PORTFOLIO_URL

_TEMPLATE_OTHER = """\
Hi {{name}},

I came across the {job_title} opening at {company_name} and I'm very interested in the role.

I recently completed my M.S. in Information Systems (SAP) at CMU in December 2025 with a 3.95 GPA, and I have 5+ years of software engineering experience with a growing specialization in AI/ML systems including RAG pipelines, LLM integration, FAISS vector search, and full-stack development with React, Next.js, Node.js, and Python (FastAPI).

Most recently, I built FinanceAI, an end-to-end AI-powered personal finance application, as a personal project to deepen my AI engineering skills. I'm currently interning at Detroit Manufacturing Systems in Detroit while actively pursuing full-time opportunities. You can view this project and more on my portfolio here: """ + PORTFOLIO_URL + """

On a lighter note, I competed in a stand-up comedy competition at CMU and secured 2nd place based on audience voting. I bring that same creativity, confidence, and outside-the-box thinking to every team I work with.

I've attached my resume for your reference and would love to connect — even a quick 15-minute conversation would be incredibly valuable.

Thank you so much for your time, {{name}}!

Best,
Mayank Surti
(248) 704-9118 | linkedin.com/in/mayank-surti-593bb3185/
github.com/surti1mr
Portfolio: """ + PORTFOLIO_URL

_TEMPLATES = {
    "cmu_alumni":  _TEMPLATE_CMU_ALUMNI,
    "recruiter":   _TEMPLATE_RECRUITER,
    "it_employee": _TEMPLATE_IT_EMPLOYEE,
    "other":       _TEMPLATE_OTHER,
}

# ── Groq only extracts job_title, company_name, and subject ──────────────────

_EXTRACT_SYSTEM = (
    "You are a precise information extractor. "
    "Return ONLY valid JSON with exactly the keys requested. No extra text."
)

_EXTRACT_PROMPT = """\
From the job description below, extract:
1. "job_title"   — the exact job title listed (e.g. "Software Engineer")
2. "company_name" — the exact company name (e.g. "Ford Motor Company")
3. "subject"     — a cold email subject line: under 12 words, plain language, no exclamation mark

Return ONLY this JSON (no markdown, no explanation):
{{"job_title": "...", "company_name": "...", "subject": "..."}}

===JOB DESCRIPTION===
{job_description}"""


def generate_email(
    resume_text: str,
    job_description: str,
    recipient_type: str,
    custom_position: str = "",
) -> dict[str, str]:
    """Returns {"subject": "...", "body": "..."} or raises RuntimeError."""
    api_key = os.getenv("GROQ_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("GROQ_API_KEY not set in .env file.")

    from groq import Groq

    client = Groq(api_key=api_key)

    # Step 1: extract job_title, company_name, subject from the JD
    extract_resp = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": _EXTRACT_SYSTEM},
            {"role": "user",   "content": _EXTRACT_PROMPT.format(job_description=job_description.strip())},
        ],
        temperature=0.0,
        max_tokens=120,
    )

    raw_json = extract_resp.choices[0].message.content.strip()
    extracted = _parse_json(raw_json)

    job_title    = extracted.get("job_title",    "Software Engineer").strip()
    company_name = extracted.get("company_name", "your company").strip()

    if recipient_type == "cmu_alumni":
        subject = f"CMU Alum \u2013 Exploring {job_title} Opportunity at {company_name}"
    elif recipient_type == "it_employee":
        subject = f"Interested in {job_title} at {company_name} \u2013 Would Love Your Perspective"
    elif recipient_type == "recruiter":
        subject = f"Strong Fit for {job_title} \u2013 AI/Full-Stack Engineer | 5+ Years Experience"
    else:
        subject = extracted.get("subject", f"{job_title} at {company_name} \u2014 Mayank Surti").strip()

    # Step 2: fill the hardcoded template
    template = _TEMPLATES.get(recipient_type, _TEMPLATE_OTHER)
    body = template.format(job_title=job_title, company_name=company_name)

    return {"subject": subject, "body": body, "company_name": company_name, "job_title": job_title}


def _parse_json(raw: str) -> dict:
    # Strip markdown code fences if present
    raw = re.sub(r"^```[a-z]*\n?", "", raw.strip())
    raw = re.sub(r"\n?```$", "", raw.strip())
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # Try to extract keys manually as fallback
        result = {}
        for key in ("job_title", "company_name", "subject"):
            m = re.search(rf'"{key}"\s*:\s*"([^"]+)"', raw)
            if m:
                result[key] = m.group(1)
        return result
