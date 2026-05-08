"""
Cold Mailer — Flask Backend
Run: python3 app.py
Open: http://<your-ec2-ip>:5000
"""

import os, json, smtplib, ssl
from flask import Flask, request, jsonify, render_template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import openpyxl
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "coldmailer_shivam_2024"

UPLOAD_FOLDER = "/tmp/cold_mailer_uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".cold_mailer_config.json")

SUBJECTS = [
    "Aspiring DevOps Engineer – Open to Opportunities",
    "DevOps & Cloud Engineer – Seeking Role at {company}",
    "Fresher DevOps Engineer – AWS & Kubernetes Experience",
    "Application for DevOps/Cloud Engineer Position",
    "Shivam Garud – DevOps Engineer | AWS | Docker | K8s",
]

# ── CONFIG ────────────────────────────────────────────────────────────────────
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE) as f:
            return json.load(f)
    return {"name": "Shivam Garud", "email": "", "app_password": ""}

def save_config_file(cfg):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f)

# ── EMAIL BODY ────────────────────────────────────────────────────────────────
def generate_body(name, title, company):
    if title and title.strip() and title.lower() not in ("none", "nan", "-", ""):
        greeting = f"Dear {title} {name}"
    else:
        greeting = "Dear Hiring Manager"
    return f"""{greeting},

I hope you're doing well.

I'm Shivam Garud, an aspiring DevOps Engineer with hands-on experience in AWS (EC2, S3, VPC, IAM, Lambda, Route 53, Auto Scaling, Load Balancers, RDS, EFS) along with tools like Docker, Kubernetes, Terraform, Git, and monitoring using Grafana & Prometheus.

I'm a quick learner, adaptable, and passionate about automation and building efficient systems. I believe my skills align well with the kind of work done at {company}.

Please find my work below:
Portfolio  : https://shivam-garud.vercel.app/
LinkedIn   : https://www.linkedin.com/in/shivam-garud/
GitHub     : https://github.com/Shivamgarud8

I've also attached my resume for your review.

Looking forward to your response.

Best regards,
Shivam Garud
+91 7434895001"""

# ── ROUTES ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/config", methods=["GET"])
def get_config():
    cfg = load_config()
    cfg.pop("app_password", None)   # never send password to browser
    return jsonify(cfg)

@app.route("/api/config", methods=["POST"])
def post_config():
    data = request.json
    cfg = load_config()
    for k in ("name", "email", "app_password"):
        if k in data and data[k]:
            cfg[k] = data[k]
    save_config_file(cfg)
    return jsonify({"ok": True})

@app.route("/api/upload_excel", methods=["POST"])
def upload_excel():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400
    path = os.path.join(UPLOAD_FOLDER, "data.xlsx")
    f.save(path)
    try:
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return jsonify({"error": "Excel file is empty"}), 400

        # Auto-detect columns by header keywords
        header = [str(h).strip().lower() if h else "" for h in rows[0]]
        kws = {
            "sr":      ["sr", "no", "number", "serial", "id", "#"],
            "name":    ["name"],
            "email":   ["email", "mail", "e-mail"],
            "title":   ["title", "designation", "role", "position"],
            "company": ["company", "org", "organisation", "organization", "firm"],
        }
        col_map = {}
        for field, keywords in kws.items():
            for i, h in enumerate(header):
                if any(k in h for k in keywords):
                    col_map[field] = i
                    break

        data = []
        for idx, r in enumerate(rows[1:], 1):
            def get(field):
                if field in col_map:
                    val = r[col_map[field]]
                    return str(val).strip() if val else ""
                return ""
            em = get("email")
            if em and em.lower() not in ("none", "nan", ""):
                data.append({
                    "sr":      get("sr") or str(idx),
                    "name":    get("name"),
                    "email":   em,
                    "title":   get("title"),
                    "company": get("company"),
                })
        return jsonify({"rows": data, "total": len(data)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/upload_resume", methods=["POST"])
def upload_resume():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400
    fname = secure_filename(f.filename)
    path = os.path.join(UPLOAD_FOLDER, fname)
    f.save(path)
    return jsonify({"ok": True, "filename": fname, "path": path})

@app.route("/api/generate", methods=["POST"])
def generate():
    data = request.json
    rows = data.get("rows", [])
    subj_idx = int(data.get("subject_idx", 0))
    emails = []
    for r in rows:
        company = r.get("company") or "your company"
        subj = SUBJECTS[subj_idx].replace("{company}", company)
        body = generate_body(r.get("name", ""), r.get("title", ""), company)
        emails.append({
            "to":      r["email"],
            "name":    r.get("name", ""),
            "subject": subj,
            "body":    body,
        })
    return jsonify({"emails": emails})

@app.route("/api/send", methods=["POST"])
def send_emails():
    data = request.json
    emails = data.get("emails", [])
    resume_path = data.get("resume_path", "")
    cfg = load_config()

    if not cfg.get("email") or not cfg.get("app_password"):
        return jsonify({"error": "Gmail settings not saved. Go to Settings tab first."}), 400
    if not emails:
        return jsonify({"error": "No emails to send"}), 400

    sent, failed = 0, []
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(cfg["email"], cfg["app_password"])
            for em in emails:
                try:
                    msg = MIMEMultipart()
                    msg["From"]    = f"{cfg.get('name', 'Shivam Garud')} <{cfg['email']}>"
                    msg["To"]      = em["to"]
                    msg["Subject"] = em["subject"]
                    msg.attach(MIMEText(em["body"], "plain"))
                    # Attach resume if uploaded
                    if resume_path and os.path.exists(resume_path):
                        with open(resume_path, "rb") as rf:
                            part = MIMEApplication(rf.read(), Name=os.path.basename(resume_path))
                            part["Content-Disposition"] = f'attachment; filename="{os.path.basename(resume_path)}"'
                            msg.attach(part)
                    server.sendmail(cfg["email"], em["to"], msg.as_string())
                    sent += 1
                except Exception as e:
                    failed.append({"to": em["to"], "error": str(e)})
    except smtplib.SMTPAuthenticationError:
        return jsonify({"error": "Gmail Auth Failed! Use App Password not your regular password."}), 401
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    return jsonify({"sent": sent, "failed": failed})


if __name__ == "__main__":
    print("\n╔══════════════════════════════════════════╗")
    print("║        COLD MAILER — Shivam Garud        ║")
    print("╚══════════════════════════════════════════╝")
    print(f"\n✅  Server running!")
    print(f"🌐  Local  : http://localhost:5000")
    print(f"🌐  EC2    : http://0.0.0.0:5000")
    print(f"\n   Open port 5000 in your EC2 Security Group!\n")
    app.run(host="0.0.0.0", port=5000, debug=False)
