"""Local web UI for personalized cold emails — Playwright OWA / SMTP / Outlook COM."""

from __future__ import annotations

import os
import re
import threading
import uuid
from pathlib import Path

from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request

load_dotenv(Path(__file__).resolve().parent / ".env")

import email_backend
import groq_generator
import database
import playwright_sender
import smtp_sender
from outlook_sender import preview_first

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

DEFAULT_FROM = "surti1mr@cmich.edu"
DEFAULT_SUBJECT = "Introduction - AI Software Engineer (Mayank Surti)"
DEFAULT_BODY = """Hi {{name}},

I hope you're doing well! I recently came across the AI Software Engineer contract role in Farmington Hills through LinkedIn and wanted to introduce myself.

I'm Mayank, A software engineer based in Michigan with 5+ years of experience, recently focused on AI systems and LLM integration. I just wrapped up my M.S. in Information Systems and built FinanceAI as a personal project: a full-stack RAG application using FastAPI, FAISS, LLaMA 3.3 70B, and Next.js, with semantic search, auto-categorization, and per-user data isolation. I'm also currently interning at Detroit Manufacturing Systems in Detroit, so I'm local and immediately available.

The role's focus on guardrail design, enforcement patterns across services, and agentic workflows resonates with work I've done around phishing detection, cross-service API design, and LLM-powered pipelines. I'd love to chat and see if there's a mutual fit.

Would you have 15 minutes this week for a quick call?

Thanks so much for your time, {{name}}!

Mayank Surti
(248) 704-9118
surti1mr@cmich.edu
https://portfolio-ten-pi-50.vercel.app"""


def _form_data():
    return {
        "from_email": DEFAULT_FROM,
        "subject": request.form.get("subject", "").strip(),
        "body_template": request.form.get("body_template", ""),
        "recipients": request.form.get("recipients", "").strip(),
        "as_draft": request.form.get("as_draft") == "on",
        "delay_seconds": float(request.form.get("delay_seconds") or 5),
    }


def _save_resume() -> tuple[str, str] | None:
    f = request.files.get("resume")
    if not f or not f.filename:
        path = request.form.get("resume_path")
        if path and Path(path).is_file():
            return path, Path(path).name
        return None
    ext = Path(f.filename).suffix.lower()
    if ext not in {".pdf", ".doc", ".docx"}:
        raise ValueError("Resume must be .pdf, .doc, or .docx")
    original_name = Path(f.filename).name
    safe_name = re.sub(r'[<>:"/\\|?\x00-\x1f]', "_", original_name).strip() or f"resume{ext}"
    path = UPLOAD_DIR / f"{uuid.uuid4().hex}_{safe_name}"
    f.save(path)
    return str(path), original_name


@app.route("/")
def index():
    status = email_backend.get_status()
    return render_template(
        "index.html",
        default_from=DEFAULT_FROM,
        from_display=DEFAULT_FROM,
        default_subject=DEFAULT_SUBJECT,
        default_body=DEFAULT_BODY,
        status=status,
        db_available=database.is_available(),
    )


# --- Outlook Web sign-in -------------------------------------------------------

_login_status = {"running": False, "done": False, "error": None}


def _run_login():
    global _login_status
    _login_status = {"running": True, "done": False, "error": None}
    try:
        playwright_sender.do_browser_login()
        _login_status = {"running": False, "done": True, "error": None}
    except Exception as exc:
        _login_status = {"running": False, "done": False, "error": str(exc)}


@app.route("/auth/owa-login", methods=["POST"])
def auth_owa_login():
    if _login_status["running"]:
        return jsonify({"ok": False, "error": "Sign-in already in progress"}), 400
    t = threading.Thread(target=_run_login, daemon=True)
    t.start()
    return jsonify({"ok": True, "message": "Browser opened — sign in and click Continue."})


@app.route("/auth/owa-status")
def auth_owa_status():
    logged_in = playwright_sender.is_logged_in()
    return jsonify({
        "ok": True,
        "logged_in": logged_in,
        "running": _login_status["running"],
        "done": _login_status["done"],
        "error": _login_status["error"],
    })


@app.route("/auth/owa-logout", methods=["POST"])
def auth_owa_logout():
    playwright_sender.logout()
    return jsonify({"ok": True})


# --- AI email generation ------------------------------------------------------

@app.route("/api/generate", methods=["POST"])
def api_generate():
    try:
        resume_text = request.form.get("resume_text", "").strip()
        job_description = request.form.get("job_description", "").strip()
        recipient_type = request.form.get("recipient_type", "recruiter").strip()
        custom_position = request.form.get("custom_position", "").strip()

        if not resume_text:
            return jsonify({"ok": False, "error": "Paste your resume text first."}), 400
        if not job_description:
            return jsonify({"ok": False, "error": "Paste the job description first."}), 400

        result = groq_generator.generate_email(
            resume_text=resume_text,
            job_description=job_description,
            recipient_type=recipient_type,
            custom_position=custom_position,
        )
        return jsonify({"ok": True, "subject": result["subject"], "body": result["body"], "company_name": result.get("company_name", ""), "job_title": result.get("job_title", "")})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


# --- Preview ------------------------------------------------------------------

@app.route("/api/preview", methods=["POST"])
def api_preview():
    try:
        data = _form_data()
        preview = preview_first(
            subject=data["subject"],
            body_template=data["body_template"],
            recipients_text=data["recipients"],
            from_email=data["from_email"],
        )
        return jsonify({"ok": True, "preview": preview})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


# --- Send (auto) --------------------------------------------------------------

@app.route("/api/send", methods=["POST"])
def api_send():
    try:
        status = email_backend.get_status()
        if not status["can_auto_send"]:
            return jsonify({"ok": False, "error": status["message"]}), 400

        data = _form_data()
        saved = _save_resume()
        scheduled_at = request.form.get("scheduled_at", "").strip() or None
        results, out_dir = email_backend.send_batch(
            mode="auto",
            from_email=data["from_email"],
            subject=data["subject"],
            body_template=data["body_template"],
            recipients_text=data["recipients"],
            attachment_path=saved[0] if saved else None,
            attachment_filename=saved[1] if saved else None,
            as_draft=data["as_draft"] if status["backend"] == "com" else False,
            delay_seconds=max(0, data["delay_seconds"]),
            scheduled_at=scheduled_at,
        )
        sent = sum(1 for r in results if r.success)

        # Auto-save each successful send to MySQL tracking DB (optional — never blocks send)
        recipient_type = request.form.get("recipient_type", "").strip()
        company_name   = request.form.get("company_name",   "").strip()
        tracked = 0
        for r in results:
            if r.success and database.is_available():
                try:
                    if database.add_contact(
                        company_name=company_name,
                        person_name=r.name,
                        email=r.email,
                        type_=recipient_type,
                        status="Delivered",
                    ):
                        tracked += 1
                except Exception:
                    pass

        return jsonify({
            "ok": True,
            "sent": sent,
            "total": len(results),
            "tracked": tracked,
            "db_available": database.is_available(),
            "backend": status["backend"],
            "output_folder": out_dir,
            "results": [
                {"email": r.email, "name": r.name, "success": r.success, "message": r.message}
                for r in results
            ],
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


# --- Prepare .eml fallback ----------------------------------------------------

@app.route("/api/prepare-eml", methods=["POST"])
def api_prepare_eml():
    try:
        data = _form_data()
        saved = _save_resume()
        results, out_dir = email_backend.send_batch(
            mode="eml",
            from_email=data["from_email"],
            subject=data["subject"],
            body_template=data["body_template"],
            recipients_text=data["recipients"],
            attachment_path=saved[0] if saved else None,
            attachment_filename=saved[1] if saved else None,
            delay_seconds=max(0, data["delay_seconds"]),
            open_in_outlook=True,
        )
        sent = sum(1 for r in results if r.success)
        return jsonify({
            "ok": True,
            "sent": sent,
            "total": len(results),
            "backend": "eml",
            "output_folder": out_dir,
            "results": [
                {"email": r.email, "name": r.name, "success": r.success, "message": r.message}
                for r in results
            ],
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


# --- Open folder --------------------------------------------------------------

@app.route("/api/open-folder", methods=["POST"])
def api_open_folder():
    folder = (request.json or {}).get("folder", "")
    path = Path(folder)
    if not path.is_dir() or not str(path.resolve()).startswith(
        str((BASE_DIR / "generated_emails").resolve())
    ):
        return jsonify({"ok": False, "error": "Invalid folder"}), 400
    os.startfile(str(path))
    return jsonify({"ok": True})


# --- Tracking -----------------------------------------------------------------

@app.route("/tracking")
def tracking():
    return render_template("tracking.html", db_available=database.is_available())


@app.route("/api/contacts", methods=["GET"])
def api_contacts_list():
    try:
        rows = database.get_all_contacts()
        return jsonify({"ok": True, "contacts": rows})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.route("/api/contacts", methods=["POST"])
def api_contacts_add():
    try:
        body = request.get_json(force=True) or {}
        new_id = database.add_contact(
            company_name=body.get("company_name", ""),
            person_name=body.get("person_name", ""),
            email=body.get("email", ""),
            type_=body.get("type", ""),
            status=body.get("status", "Delivered"),
            notes=body.get("notes", ""),
        )
        return jsonify({"ok": True, "id": new_id})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.route("/api/contacts/<int:contact_id>", methods=["PUT"])
def api_contacts_update(contact_id):
    try:
        body = request.get_json(force=True) or {}
        allowed = {"company_name", "person_name", "email", "type", "status", "notes"}
        fields = {k: v for k, v in body.items() if k in allowed}
        database.update_contact(contact_id, **fields)
        return jsonify({"ok": True})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.route("/api/contacts/<int:contact_id>", methods=["DELETE"])
def api_contacts_delete(contact_id):
    try:
        database.delete_contact(contact_id)
        return jsonify({"ok": True})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


if __name__ == "__main__":
    if database.init_db():
        print("Tracking database connected.")
    else:
        print("Tracking database unavailable — emails will still send.")
        if database.last_error():
            print(f"  ({database.last_error()})")
    s = email_backend.get_status()
    print("Open http://127.0.0.1:5000 in your browser")
    print(s["message"])
    app.run(host="127.0.0.1", port=5000, debug=False)
