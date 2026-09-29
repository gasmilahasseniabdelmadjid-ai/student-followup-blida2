# PRODUCTION_VERSION = 2026-09-29-ui-fix-2
import os, io, csv, json, secrets, sqlite3, urllib.request
from datetime import datetime
from functools import wraps
from flask import Flask, request, redirect, url_for, session, send_file, abort, render_template_string, Response
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.utils import secure_filename
from PIL import Image as PILImage
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import barcode
from barcode.writer import ImageWriter
import qrcode
import arabic_reshaper
from bidi.algorithm import get_display

BASE=os.path.dirname(os.path.abspath(__file__))
DATA_DIR=os.environ.get("DATA_DIR", os.path.join(BASE,"data"))
os.makedirs(os.path.join(DATA_DIR,"receipts"), exist_ok=True)
DB=os.path.join(DATA_DIR,"survey.db")
RECEIPTS=os.path.join(DATA_DIR,"receipts")
COVERS=os.path.join(DATA_DIR,"covers")
os.makedirs(COVERS, exist_ok=True)

FONT="/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
BOLD="/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
if os.path.exists(FONT): pdfmetrics.registerFont(TTFont("DejaVu",FONT))
if os.path.exists(BOLD): pdfmetrics.registerFont(TTFont("DejaVuBold",BOLD))

app=Flask(__name__)
app.wsgi_app=ProxyFix(app.wsgi_app,x_for=1,x_proto=1,x_host=1)
app.secret_key=os.environ.get("SECRET_KEY","local-change-this")
app.config.update(SESSION_COOKIE_HTTPONLY=True,SESSION_COOKIE_SAMESITE="Lax",MAX_CONTENT_LENGTH=8*1024*1024)
if os.environ.get("COOKIE_SECURE","1")=="1": app.config["SESSION_COOKIE_SECURE"]=True
ADMIN_PASSWORD=os.environ.get("ADMIN_PASSWORD","change-me")

RANKS=["أستاذ","أستاذ محاضر أ","أستاذ محاضر ب","أستاذ مساعد أ","أستاذ مساعد ب","طالب دكتوراه"]
CLASSIFICATIONS=["A+","A","B","C","غير مصنف"]
DATABASES=["Scopus","Web of Science","ASJP","Google Scholar","DOAJ","Other","Not indexed"]

T={
"ar":{"dir":"rtl","univ":"جامعة البليدة2","vice":"نيابة رئاسة الجامعة للبحث العلمي وما بعد التدرج","platform":"منصة جمع بيانات النشر العلمي للأساتذة وطلبة الدكتوراه","title":"النشر العلمي | جامعة البليدة2","survey":"منصة جمع بيانات النشر العلمي","intro":"تتيح المنصة التصريح بالمنشورات العلمية في إرسال واحد، ثم توليد وصل PDF يتضمن البيانات ورمزاً شريطياً ورمز QR للتحقق الإلكتروني.","notice":"يرجى إدخال المعلومات بدقة، وخاصة عنوان المقال واسم المجلة والتصنيف وقواعد الفهرسة.","general":"1. المعلومات العامة","pubs":"2. المنشورات العلمية","add":"إضافة مقال","ar_name":"الاسم واللقب باللغة العربية","en_name":"الاسم واللقب باللغة الإنجليزية","rank":"الرتبة","faculty":"الكلية","department":"القسم","spec":"التخصص","lab":"عنوان المخبر الذي ينتمي له الأستاذ أو الطالب","choose":"اختر...","article":"عنوان المقال","journal":"اسم المجلة","class":"نوع التصنيف","db":"قاعدة/قواعد البيانات التي ينتمي لها المقال","year":"سنة النشر","url":"رابط المقال","scholar":"رابط Google Scholar","pub_type":"نوع المنشور","article_type":"مقال","book_type_label":"كتاب","article_title_en":"عنوان المقال باللغة الإنجليزية","book_title":"عنوان الكتاب","book_kind":"طبيعة الكتاب","scientific":"علمي","pedagogical":"بيداغوجي","publisher":"دار النشر","book_year":"سنة النشر","excerpt":"رابط المستلة","cover":"صورة واجهة الكتاب","cover_hint":"PNG/JPG/WebP — الحد الأقصى 5MB","multi":"يمكن اختيار أكثر من قاعدة.","submit":"إرسال التصريح وإنشاء الوصل PDF","clear":"مسح النموذج","success":"تم تسجيل التصريح بنجاح","success2":"تم حفظ بيانات النشر العلمي وإنشاء الوصل.","download":"تحميل الوصل PDF","verify":"التحقق الإلكتروني","code":"رمز الوصل","admin":"إدارة بيانات النشر العلمي","password":"كلمة مرور الإدارة","login":"دخول","export":"تصدير CSV","logout":"تسجيل الخروج","verified":"وصل موثق","notfound":"الرمز غير موجود","count":"عدد المنشورات","date":"تاريخ الإرسال","item":"البيان","info":"المعلومات","scan":"امسح QR للتحقق","footer":"جامعة البليدة2 — نيابة رئاسة الجامعة للبحث العلمي وما بعد التدرج","bad":"يرجى إكمال جميع البيانات المطلوبة."},
"fr":{"dir":"ltr","univ":"Université de Blida 2","vice":"Vice-Rectorat de la Recherche Scientifique et de la Formation Post-Graduée","platform":"Plateforme de collecte des publications scientifiques des enseignants et doctorants","title":"Publications scientifiques | Université de Blida 2","survey":"Formulaire de collecte des publications scientifiques","intro":"La plateforme permet de déclarer plusieurs publications en une seule soumission et génère un reçu PDF avec les données, un code-barres et un QR de vérification.","notice":"Veuillez vérifier soigneusement les titres, revues, classements et bases d’indexation.","general":"1. Informations générales","pubs":"2. Publications scientifiques","add":"Ajouter un article","ar_name":"Nom et prénom en arabe","en_name":"Nom et prénom en anglais","rank":"Grade","faculty":"Faculté","department":"Département","spec":"Spécialité","lab":"Intitulé du laboratoire d’appartenance","choose":"Choisir...","article":"Titre de l’article","journal":"Nom de la revue","class":"Classement","db":"Base(s) d’indexation","year":"Année de publication","url":"Lien de l’article","scholar":"Lien Google Scholar","pub_type":"Type de publication","article_type":"Article","book_type_label":"Livre","article_title_en":"Titre de l’article en anglais","book_title":"Titre du livre","book_kind":"Nature du livre","scientific":"Scientifique","pedagogical":"Pédagogique","publisher":"Maison d’édition","book_year":"Année de publication","excerpt":"Lien de l’extrait","cover":"Image de couverture","cover_hint":"PNG/JPG/WebP — maximum 5 Mo","multi":"Plusieurs bases peuvent être sélectionnées.","submit":"Envoyer et générer le reçu PDF","clear":"Effacer","success":"Déclaration enregistrée","success2":"Les données ont été enregistrées et le reçu a été généré.","download":"Télécharger le reçu PDF","verify":"Vérification","code":"Code du reçu","admin":"Administration des publications","password":"Mot de passe administrateur","login":"Connexion","export":"Exporter CSV","logout":"Déconnexion","verified":"Reçu vérifié","notfound":"Code introuvable","count":"Nombre de publications","date":"Date d’envoi","item":"Élément","info":"Informations","scan":"Scanner le QR pour vérifier","footer":"Université de Blida 2 — Vice-Rectorat de la Recherche Scientifique et de la Formation Post-Graduée","bad":"Veuillez compléter tous les champs obligatoires."},
"en":{"dir":"ltr","univ":"University of Blida 2","vice":"Vice-Rectorate for Scientific Research and Postgraduate Studies","platform":"Scientific publication data collection platform for faculty members and PhD students","title":"Scientific Publications | University of Blida 2","survey":"Scientific Publication Data Collection Form","intro":"The platform accepts multiple publications in one submission and generates a PDF receipt containing the submitted data, a barcode and a verification QR code.","notice":"Please verify article titles, journal names, classifications and indexing databases.","general":"1. General information","pubs":"2. Scientific publications","add":"Add article","ar_name":"Full name in Arabic","en_name":"Full name in English","rank":"Academic rank","faculty":"Faculty","department":"Department","spec":"Specialization","lab":"Name of affiliated laboratory","choose":"Choose...","article":"Article title","journal":"Journal name","class":"Classification","db":"Indexing database(s)","year":"Publication year","url":"Article link","scholar":"Google Scholar link","pub_type":"Publication type","article_type":"Article","book_type_label":"Book","article_title_en":"Article title in English","book_title":"Book title","book_kind":"Book type","scientific":"Scientific","pedagogical":"Pedagogical","publisher":"Publisher","book_year":"Publication year","excerpt":"Excerpt link","cover":"Book cover image","cover_hint":"PNG/JPG/WebP — maximum 5 MB","multi":"Multiple databases may be selected.","submit":"Submit and generate PDF receipt","clear":"Clear","success":"Submission recorded","success2":"The publication data were saved and the receipt was generated.","download":"Download PDF receipt","verify":"Online verification","code":"Receipt code","admin":"Scientific Publications Administration","password":"Administrator password","login":"Login","export":"Export CSV","logout":"Logout","verified":"Verified receipt","notfound":"Code not found","count":"Number of publications","date":"Submission date","item":"Item","info":"Information","scan":"Scan QR to verify","footer":"University of Blida 2 — Vice-Rectorate for Scientific Research and Postgraduate Studies","bad":"Please complete all required fields."}}

def lang():
    x=request.args.get("lang")
    if x in T: session["lang"]=x
    return session.get("lang","ar")
def tr(k): return T[lang()][k]
def ar(s):
    try:return get_display(arabic_reshaper.reshape(str(s)))
    except Exception:return str(s)
def db():
    c=sqlite3.connect(DB,timeout=30); c.row_factory=sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON"); c.execute("PRAGMA busy_timeout=30000"); c.execute("PRAGMA journal_mode=WAL"); return c
def init_db():
    c=db(); c.executescript("""CREATE TABLE IF NOT EXISTS submissions(
    id INTEGER PRIMARY KEY AUTOINCREMENT,receipt_code TEXT UNIQUE NOT NULL,arabic_name TEXT NOT NULL,english_name TEXT NOT NULL,rank TEXT NOT NULL,
    faculty TEXT NOT NULL,department TEXT NOT NULL,specialization TEXT NOT NULL,laboratory TEXT NOT NULL,submitted_at TEXT NOT NULL,language TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS publications(
    id INTEGER PRIMARY KEY AUTOINCREMENT,submission_id INTEGER NOT NULL,article_title TEXT NOT NULL,journal_name TEXT NOT NULL,classification TEXT NOT NULL,
    databases TEXT NOT NULL,publication_year INTEGER NOT NULL,article_url TEXT,scholar_url TEXT,FOREIGN KEY(submission_id) REFERENCES submissions(id) ON DELETE CASCADE);
    CREATE INDEX IF NOT EXISTS idx_receipt ON submissions(receipt_code);
    CREATE INDEX IF NOT EXISTS idx_pub_submission ON publications(submission_id);""")
    existing={row[1] for row in c.execute("PRAGMA table_info(publications)").fetchall()}
    migrations={"publication_type":"TEXT NOT NULL DEFAULT 'article'","article_title_en":"TEXT","book_title":"TEXT","book_type":"TEXT","publisher":"TEXT","book_year":"INTEGER","excerpt_url":"TEXT","cover_path":"TEXT"}
    for name,definition in migrations.items():
        if name not in existing: c.execute(f"ALTER TABLE publications ADD COLUMN {name} {definition}")
    c.commit(); c.close()
init_db()

def csrf():
    if "csrf" not in session: session["csrf"]=secrets.token_urlsafe(32)
    return session["csrf"]
def check_csrf():
    if not request.form.get("csrf_token") or not secrets.compare_digest(request.form.get("csrf_token",""),session.get("csrf","")): abort(400,"Invalid CSRF token")
def admin_required(f):
    @wraps(f)
    def w(*a,**kw):
        if not session.get("admin"): return redirect(url_for("admin"))
        return f(*a,**kw)
    return w
@app.after_request
def headers(r):
    r.headers["X-Content-Type-Options"]="nosniff"; r.headers["X-Frame-Options"]="SAMEORIGIN"; r.headers["Referrer-Policy"]="strict-origin-when-cross-origin"
    r.headers["Permissions-Policy"]="camera=(), microphone=(), geolocation=()"
    if request.is_secure:r.headers["Strict-Transport-Security"]="max-age=31536000; includeSubDomains"
    return r

def layout(body,title=None):
    l=lang(); return render_template_string("""<!doctype html><html lang="{{l}}" dir="{{d}}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{{title}}</title>
<style>
:root{--g:#166534;--g2:#064e3b;--bg:#f5f7f6;--line:#d1d5db;--txt:#17202a}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--txt);font-family:Tahoma,"Segoe UI",Arial,sans-serif;line-height:1.65}.hero{background:linear-gradient(135deg,var(--g2),var(--g));color:white;padding:22px 16px}.top{max-width:1100px;margin:auto}.langs{display:flex;gap:6px;justify-content:flex-end;margin-bottom:12px}.langs a{background:white;color:#14532d;padding:5px 9px;border-radius:8px;text-decoration:none;font-size:13px}.brand{display:flex;align-items:center;gap:16px}.univ-logo{width:150px;height:150px;object-fit:contain;display:block;flex:none;filter:drop-shadow(0 2px 5px rgba(0,0,0,.22))}.container{max-width:1100px;margin:25px auto;padding:0 14px}.card{background:white;border:1px solid var(--line);border-radius:16px;padding:22px;margin-bottom:18px;box-shadow:0 4px 18px #0000000b}.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}label{display:flex;flex-direction:column;gap:5px;font-weight:700}label.full{grid-column:1/-1}input,select,textarea{width:100%;padding:11px;border:1px solid #cbd5e1;border-radius:9px;font:inherit;background:white}select[multiple]{min-height:110px}.hint{font-size:12px;color:#64748b;font-weight:400}.notice{background:#f0fdf4;border-inline-start:4px solid #16a34a;padding:12px;border-radius:8px}.head{display:flex;justify-content:space-between;align-items:center;gap:10px}.actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:20px}button,.btn{border:0;border-radius:9px;padding:11px 17px;font:inherit;cursor:pointer;text-decoration:none;display:inline-block}.primary{background:#166534;color:#fff}.secondary{background:#e8f5e9;color:#14532d;border:1px solid #bbd7c0}.ghost{background:#f1f5f9;color:#334155}.danger{background:#fee2e2;color:#991b1b}.article{background:#fbfffc;border:1px solid #dbe7df;border-radius:12px;padding:16px;margin-top:14px}.success{text-align:center}.check{margin:auto;width:58px;height:58px;border-radius:50%;display:grid;place-items:center;background:#dcfce7;color:#166534;font-size:32px}.code{font-size:22px;font-weight:800;letter-spacing:1px;background:#f1f5f9;padding:12px;border-radius:10px;margin:15px 0}.verify{display:grid;grid-template-columns:1fr 1fr;gap:10px;background:#f8fafc;padding:15px;border-radius:10px}.muted{color:#64748b}.error{color:#991b1b;font-weight:800}table{width:100%;border-collapse:collapse;font-size:13px}th,td{border:1px solid #d1d5db;padding:8px;text-align:start}th{background:#ecfdf5}footer{text-align:center;color:#64748b;padding:25px}.small{font-size:12px}@media(max-width:720px){.grid,.verify{grid-template-columns:1fr}.univ-logo{width:86px;height:86px}.brand h1{font-size:19px}.brand p{font-size:13px}table{display:block;overflow:auto;white-space:nowrap}}
</style></head><body>{{body|safe}}</body></html>""",l=l,d=T[l]["dir"],title=title or T[l]["title"],csrf=csrf,body=body)

def page_header():
    return f"""<header class="hero"><div class="top"><div class="langs"><a href="?lang=ar">العربية</a><a href="?lang=fr">Français</a><a href="?lang=en">English</a></div>
<div class="brand"><img class="univ-logo" src="https://elearning.univ-blida2.dz/pluginfile.php?file=%2F1%2Ftheme_academi%2Flogo%2F1766305366%2Flogo.png" alt="جامعة البليدة2"><div><h1>{tr("univ")}</h1><p>{tr("vice")}</p><p>{tr("platform")}</p></div></div></div></header>"""
def get_submission(code):
    c=db(); s=c.execute("SELECT * FROM submissions WHERE receipt_code=?",(code,)).fetchone()
    if not s:c.close();return None,[]
    p=c.execute("SELECT * FROM publications WHERE submission_id=? ORDER BY id",(s["id"],)).fetchall(); c.close(); return s,p
def code_new():
    while True:
        x="BL2-"+datetime.now().strftime("%Y")+"-"+secrets.token_hex(4).upper()
        c=db(); ok=not c.execute("SELECT 1 FROM submissions WHERE receipt_code=?",(x,)).fetchone(); c.close()
        if ok:return x

def pdf(code):
    s,p=get_submission(code)
    if not s: abort(404)
    path=os.path.join(RECEIPTS,code+".pdf"); bcpath=os.path.join(RECEIPTS,code)
    try:
        barcode.get("code128",code,writer=ImageWriter()).save(bcpath,options={"write_text":True,"module_height":13,"module_width":0.28,"font_size":9,"quiet_zone":3})
        bcp=bcpath+".png"
        verify_url=request.host_url.rstrip("/")+url_for("verify",code=code)
        qrp=os.path.join(RECEIPTS,code+"_qr.png"); qrcode.make(verify_url).save(qrp)
    except Exception:
        bcp=qrp=None
    l=s["language"] if s["language"] in T else "ar"; rtl=(l=="ar")
    tx=lambda x:ar(x) if rtl else str(x)
    st=getSampleStyleSheet()
    base=ParagraphStyle("base",parent=st["Normal"],fontName="DejaVu",fontSize=8.5,leading=13,alignment=TA_RIGHT if rtl else TA_LEFT)
    center=ParagraphStyle("center",parent=base,alignment=TA_CENTER)
    title=ParagraphStyle("title",parent=base,fontName="DejaVuBold",fontSize=15,leading=20,alignment=TA_CENTER,textColor=colors.HexColor("#14532d"))
    doc=SimpleDocTemplate(path,pagesize=A4,rightMargin=35,leftMargin=35,topMargin=28,bottomMargin=32); story=[]
    try:
        logo_data=urllib.request.urlopen("https://elearning.univ-blida2.dz/pluginfile.php?file=%2F1%2Ftheme_academi%2Flogo%2F1766305366%2Flogo.png",timeout=8).read()
        story += [Image(io.BytesIO(logo_data),width=82,height=82),Spacer(1,3)]
    except Exception: pass
    receipt_title={"ar":"وصل التصريح بالنشر العلمي","fr":"Reçu de déclaration de publication scientifique","en":"Scientific Publication Declaration Receipt"}[l]
    story += [Paragraph(tx(T[l]["univ"]),title),Paragraph(tx(T[l]["vice"]),center),Spacer(1,6),Paragraph(tx(receipt_title),title),Paragraph(tx(f"{T[l]['code']}: {code}"),center),Paragraph(tx(s["submitted_at"]),center),Spacer(1,9)]
    def make_table(rows):
        if rtl: rows=[[b,a] for a,b in rows]; widths=[365,145]
        else: widths=[145,365]
        t=Table(rows,colWidths=widths,hAlign="RIGHT" if rtl else "LEFT")
        t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),.5,colors.HexColor("#cbd5e1")),("BACKGROUND",(0,0),(-1,0),colors.HexColor("#e8f5e9")),("VALIGN",(0,0),(-1,-1),"TOP"),("ALIGN",(0,0),(-1,-1),"RIGHT" if rtl else "LEFT"),("PADDING",(0,0),(-1,-1),5)]))
        return t
    rows=[[Paragraph(tx(T[l]["item"]),base),Paragraph(tx(T[l]["info"]),base)]]
    for label,key in [("ar_name","arabic_name"),("en_name","english_name"),("rank","rank"),("faculty","faculty"),("department","department"),("spec","specialization"),("lab","laboratory")]:
        rows.append([Paragraph(tx(T[l][label]),base),Paragraph(tx(s[key]),base)])
    story += [make_table(rows),Spacer(1,12)]
    for i,x in enumerate(p,1):
        typ=x["publication_type"] or "article"
        if typ=="book":
            rows=[[Paragraph(tx(T[l]["item"]),base),Paragraph(tx(T[l]["info"]),base)],[Paragraph(tx(T[l]["book_title"]),base),Paragraph(tx(x["book_title"] or "—"),base)],[Paragraph(tx(T[l]["book_kind"]),base),Paragraph(tx(x["book_type"] or "—"),base)],[Paragraph(tx(T[l]["publisher"]),base),Paragraph(tx(x["publisher"] or "—"),base)],[Paragraph(tx(T[l]["book_year"]),base),Paragraph(str(x["book_year"] or "—"),base)],[Paragraph(tx(T[l]["excerpt"]),base),Paragraph(tx(x["excerpt_url"] or "—"),base)]]
            story += [Paragraph(tx(f"{T[l]['book_type_label']} — {i}"),title),make_table(rows)]
            if x["cover_path"] and os.path.exists(x["cover_path"]):
                try: story += [Spacer(1,5),Paragraph(tx(T[l]["cover"]),center),Image(x["cover_path"],width=125,height=155,preserveAspectRatio=True)]
                except Exception: pass
            story += [Spacer(1,10)]
        else:
            rows=[[Paragraph(tx(T[l]["item"]),base),Paragraph(tx(T[l]["info"]),base)],[Paragraph(tx(T[l]["article"]),base),Paragraph(tx(x["article_title"] or "—"),base)],[Paragraph(tx(T[l]["article_title_en"]),base),Paragraph(tx(x["article_title_en"] or "—"),base)],[Paragraph(tx(T[l]["journal"]),base),Paragraph(tx(x["journal_name"] or "—"),base)],[Paragraph(tx(T[l]["class"]),base),Paragraph(tx(x["classification"] or "—"),base)],[Paragraph(tx(T[l]["db"]),base),Paragraph(tx(x["databases"] or "—"),base)],[Paragraph(tx(T[l]["year"]),base),Paragraph(str(x["publication_year"] or "—"),base)],[Paragraph(tx(T[l]["url"]),base),Paragraph(tx(x["article_url"] or "—"),base)],[Paragraph(tx(T[l]["scholar"]),base),Paragraph(tx(x["scholar_url"] or "—"),base)]]
            story += [Paragraph(tx(f"{T[l]['article_type']} — {i}"),title),make_table(rows),Spacer(1,10)]
    if bcp and os.path.exists(bcp): story += [Spacer(1,3),Image(bcp,width=360,height=65),Paragraph(tx(code),center)]
    if qrp and os.path.exists(qrp): story += [Spacer(1,3),Image(qrp,width=82,height=82),Paragraph(tx(T[l]["scan"]),center)]
    doc.build(story); return path

@app.get("/healthz")
def health(): return {"status":"ok","service":"blida2-scientific-publications"}

@app.get("/")
def index():
    body=page_header()+f"""<main class="container"><section class="card"><h2>{tr("platform")}</h2><p>{tr("intro")}</p><div class="notice">{tr("notice")}</div></section>
<form class="card" method="post" action="/submit" id="f" enctype="multipart/form-data"><input type="hidden" name="csrf_token" value="{csrf()}">
<h2>{tr("general")}</h2><div class="grid">
<label>{tr("ar_name")} *<input name="arabic_name" required></label><label>{tr("en_name")} *<input name="english_name" required></label>
<label>{tr("rank")} *<select name="rank" required><option value="">{tr("choose")}</option>{''.join(f'<option>{x}</option>' for x in RANKS)}</select></label>
<label>{tr("faculty")} *<input name="faculty" required></label><label>{tr("department")} *<input name="department" required></label><label>{tr("spec")} *<input name="specialization" required></label>
<label class="full">{tr("lab")} *<input name="laboratory" required></label></div>
<div class="head"><h2>{tr("pubs")}</h2><button type="button" class="secondary" onclick="addPublication()">+ {tr("add")}</button></div><div id="publications"></div>
<div class="actions"><button class="primary" type="submit">{tr("submit")}</button><button class="ghost" type="button" onclick="location.reload()">{tr("clear")}</button></div></form></main><footer>{tr("footer")}</footer>
<template id="tpl"><article class="article"><div class="head"><h3>#<span class="n"></span></h3><button type="button" class="danger" onclick="this.closest('.article').remove();renumber()">×</button></div>
<label>{tr("pub_type")} *<select class="ptype" name="publication_type[]" required onchange="togglePublication(this)"><option value="article">{tr("article_type")}</option><option value="book">{tr("book_type_label")}</option></select></label>
<div class="article-fields"><div class="grid">
<label class="full">{tr("article")} *<textarea name="article_title[]" rows="2" required></textarea></label>
<label class="full">{tr("article_title_en")} *<textarea name="article_title_en[]" rows="2" required></textarea></label>
<label>{tr("journal")} *<input name="journal_name[]" required></label><label>{tr("class")} *<select name="classification[]" required>{''.join(f'<option>{x}</option>' for x in CLASSIFICATIONS)}</select></label>
<label>{tr("year")} *<input type="number" name="publication_year[]" min="1900" max="{datetime.now().year+1}" required></label>
<label class="full">{tr("db")}<select class="db" multiple size="4">{''.join(f'<option value="{x}">{x}</option>' for x in DATABASES)}</select><input type="hidden" class="dbj" name="databases_json[]"><span class="hint">{tr("multi")}</span></label>
<label>{tr("url")}<input type="url" name="article_url[]" placeholder="https://..."></label><label>{tr("scholar")}<input type="url" name="scholar_url[]" placeholder="https://scholar.google.com/..."></label></div></div>
<div class="book-fields" style="display:none"><div class="grid">
<label class="full">{tr("book_title")} *<input name="book_title[]" disabled></label><label>{tr("book_kind")} *<select name="book_type[]" disabled><option value="">{tr("choose")}</option><option value="علمي">{tr("scientific")}</option><option value="بيداغوجي">{tr("pedagogical")}</option></select></label>
<label>{tr("publisher")} *<input name="publisher[]" disabled></label><label>{tr("book_year")} *<input type="number" name="book_year[]" min="1900" max="{datetime.now().year+1}" disabled></label>
<label class="full">{tr("excerpt")}<input type="url" name="excerpt_url[]" placeholder="https://..." disabled></label>
<label class="full">{tr("cover")} *<input class="cover" type="file" name="cover[]" accept="image/png,image/jpeg,image/webp" disabled><span class="hint">{tr("cover_hint")}</span></label>
</div></div></article></template>
<script>
const box=document.getElementById('publications'),tpl=document.getElementById('tpl'),form=document.getElementById('f');
function addPublication(){box.appendChild(tpl.content.cloneNode(true));renumber()}
function renumber(){[...box.children].forEach((x,i)=>x.querySelector('.n').textContent=i+1)}
function togglePublication(sel){const card=sel.closest('.article'),isBook=sel.value==='book';card.querySelector('.article-fields').style.display=isBook?'none':'block';card.querySelector('.book-fields').style.display=isBook?'block':'none';card.querySelectorAll('.article-fields input,.article-fields textarea,.article-fields select').forEach(x=>x.disabled=isBook);card.querySelectorAll('.book-fields input:not([type="file"]),.book-fields select').forEach(x=>x.disabled=!isBook);card.querySelectorAll('.book-fields input[type="file"]').forEach(x=>x.disabled=!isBook);card.querySelectorAll('.article-fields [required]').forEach(x=>x.required=!isBook);card.querySelectorAll('.book-fields [required]').forEach(x=>x.required=isBook)}
form.addEventListener('submit',()=>box.querySelectorAll('.article').forEach(a=>a.querySelector('.dbj').value=JSON.stringify([...a.querySelectorAll('.db option:checked')].map(x=>x.value))));
addPublication();
</script>"""
    return layout(body)

@app.post("/submit")
def submit():
    check_csrf()
    vals={k:request.form.get(k,"").strip() for k in ["arabic_name","english_name","rank","faculty","department","specialization","laboratory"]}
    if not all(vals.values()) or vals["rank"] not in RANKS:return redirect(url_for("index",lang=lang()))
    types=request.form.getlist("publication_type[]")
    titles=request.form.getlist("article_title[]"); titles_en=request.form.getlist("article_title_en[]"); journals=request.form.getlist("journal_name[]"); classes=request.form.getlist("classification[]"); years=request.form.getlist("publication_year[]"); urls=request.form.getlist("article_url[]"); scholars=request.form.getlist("scholar_url[]"); dbjs=request.form.getlist("databases_json[]")
    book_titles=request.form.getlist("book_title[]"); book_types=request.form.getlist("book_type[]"); publishers=request.form.getlist("publisher[]"); book_years=request.form.getlist("book_year[]"); excerpt_urls=request.form.getlist("excerpt_url[]"); covers=request.files.getlist("cover[]")
    if not types:return redirect(url_for("index",lang=lang()))
    code=code_new(); now=datetime.now().strftime("%Y-%m-%d %H:%M:%S"); c=db(); saved_covers=[]
    try:
        sid=c.execute("INSERT INTO submissions(receipt_code,arabic_name,english_name,rank,faculty,department,specialization,laboratory,submitted_at,language) VALUES(?,?,?,?,?,?,?,?,?,?)",(code,*[vals[x] for x in ["arabic_name","english_name","rank","faculty","department","specialization","laboratory"]],now,lang())).lastrowid
        for i,typ in enumerate(types):
            if typ=="article":
                title=titles[i].strip() if i<len(titles) else ""; title_en=titles_en[i].strip() if i<len(titles_en) else ""; journal=journals[i].strip() if i<len(journals) else ""; cls=classes[i].strip() if i<len(classes) else ""; yi=int(years[i]) if i<len(years) and years[i] else 0
                if not title or not title_en or not journal or cls not in CLASSIFICATIONS or yi<1900 or yi>datetime.now().year+1:raise ValueError()
                try:selected=json.loads(dbjs[i])
                except Exception:selected=[]
                selected=[x for x in selected if x in DATABASES] or ["Not indexed"]
                c.execute("INSERT INTO publications(submission_id,publication_type,article_title,article_title_en,journal_name,classification,databases,publication_year,article_url,scholar_url) VALUES(?,?,?,?,?,?,?,?,?,?)",(sid,"article",title,title_en,journal,cls,",".join(selected),yi,urls[i].strip() if i<len(urls) else "",scholars[i].strip() if i<len(scholars) else ""))
            elif typ=="book":
                bt=book_titles[i].strip() if i<len(book_titles) else ""; bk=book_types[i].strip() if i<len(book_types) else ""; pub=publishers[i].strip() if i<len(publishers) else ""; by=int(book_years[i]) if i<len(book_years) and book_years[i] else 0
                if not bt or bk not in ("علمي","بيداغوجي") or not pub or by<1900 or by>datetime.now().year+1:raise ValueError()
                f=covers[i] if i<len(covers) else None
                if not f or not f.filename:raise ValueError()
                ext=os.path.splitext(secure_filename(f.filename))[1].lower()
                if ext not in (".png",".jpg",".jpeg",".webp"):raise ValueError()
                f.stream.seek(0); im=PILImage.open(f.stream); im.verify(); f.stream.seek(0)
                if (f.content_length or 0)>5*1024*1024:raise ValueError()
                fname=secrets.token_hex(16)+ext; cover_path=os.path.join(COVERS,fname); f.save(cover_path); saved_covers.append(cover_path)
                c.execute("INSERT INTO publications(submission_id,publication_type,article_title,article_title_en,journal_name,classification,databases,publication_year,article_url,scholar_url,book_title,book_type,publisher,book_year,excerpt_url,cover_path) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,"book",bt,"","","","",0,"","",bt,bk,pub,by,excerpt_urls[i].strip() if i<len(excerpt_urls) else "",cover_path))
            else:raise ValueError()
        c.commit()
    except Exception:
        c.rollback();c.close()
        for fp in saved_covers:
            try:os.remove(fp)
            except Exception:pass
        return redirect(url_for("index",lang=lang()))
    c.close();pdf(code);return redirect(url_for("success",code=code,lang=lang()))

@app.get("/success/<code>")
def success(code):
    s,p=get_submission(code)
    if not s:abort(404)
    body=page_header()+f"""<main class="container"><section class="card success"><div class="check">✓</div><h1>{tr("success")}</h1><p>{tr("success2")}</p><div class="code">{code}</div><a class="primary btn" href="{url_for("receipt",code=code)}">{tr("download")}</a><a class="secondary btn" href="{url_for("verify",code=code,lang=lang())}">{tr("verify")}</a></section></main><footer>{tr("footer")}</footer>"""
    return layout(body)

@app.get("/receipt/<code>.pdf")
def receipt(code):
    return send_file(pdf(code),as_attachment=True,download_name=f"Blida2_Publication_Receipt_{code}.pdf")

@app.get("/verify/<code>")
def verify(code):
    s,p=get_submission(code)
    if not s: body=f'<main class="container"><section class="card"><h1>{tr("notfound")}</h1><p>{tr("code")}: {code}</p></section></main>'
    else:
        body=page_header()+f"""<main class="container"><section class="card"><h1>✓ {tr("verified")}</h1><div class="code">{code}</div><div class="verify"><div><b>{tr("ar_name")}:</b> {s["arabic_name"]}</div><div><b>{tr("rank")}:</b> {s["rank"]}</div><div><b>{tr("faculty")}:</b> {s["faculty"]}</div><div><b>{tr("department")}:</b> {s["department"]}</div><div><b>{tr("count")}:</b> {len(p)}</div><div><b>{tr("date")}:</b> {s["submitted_at"]}</div></div></section></main><footer>{tr("footer")}</footer>"""
    return layout(body)

@app.route("/admin",methods=["GET","POST"])
def admin():
    if request.method=="POST":
        check_csrf()
        if secrets.compare_digest(request.form.get("password",""),ADMIN_PASSWORD):session["admin"]=True;return redirect(url_for("admin",lang=lang()))
    if not session.get("admin"):
        body=page_header()+f'<main class="container"><section class="card"><h1>{tr("admin")}</h1><form method="post"><input type="hidden" name="csrf_token" value="{csrf()}"><label>{tr("password")}<input type="password" name="password" required></label><button class="primary" type="submit">{tr("login")}</button></form></section></main>'
        return layout(body)
    c=db();rows=c.execute("SELECT s.*,COUNT(p.id) publication_count FROM submissions s LEFT JOIN publications p ON p.submission_id=s.id GROUP BY s.id ORDER BY s.id DESC").fetchall();c.close()
    trs="".join(f'<tr><td>{r["receipt_code"]}</td><td>{r["arabic_name"]}</td><td>{r["rank"]}</td><td>{r["faculty"]}</td><td>{r["department"]}</td><td>{r["publication_count"]}</td><td>{r["submitted_at"]}</td><td><a href="{url_for("receipt",code=r["receipt_code"])}">PDF</a></td></tr>' for r in rows)
    body=page_header()+f'<main class="container"><section class="card"><div class="head"><h1>{tr("admin")}</h1><a class="secondary btn" href="{url_for("export_csv")}">{tr("export")}</a></div><div style="overflow:auto"><table><tr><th>{tr("code")}</th><th>{tr("ar_name")}</th><th>{tr("rank")}</th><th>{tr("faculty")}</th><th>{tr("department")}</th><th>{tr("count")}</th><th>{tr("date")}</th><th>PDF</th></tr>{trs}</table></div><form method="post" action="{url_for("logout")}"><input type="hidden" name="csrf_token" value="{csrf()}"><button class="ghost" type="submit">{tr("logout")}</button></form></section></main>'
    return layout(body)

@app.get("/admin/export.csv")
@admin_required
def export_csv():
    c=db();rows=c.execute("SELECT s.receipt_code,s.arabic_name,s.english_name,s.rank,s.faculty,s.department,s.specialization,s.laboratory,s.submitted_at,p.publication_type,p.article_title,p.article_title_en,p.journal_name,p.classification,p.databases,p.publication_year,p.article_url,p.scholar_url,p.book_title,p.book_type,p.publisher,p.book_year,p.excerpt_url,p.cover_path FROM submissions s JOIN publications p ON p.submission_id=s.id ORDER BY s.id DESC,p.id").fetchall();c.close()
    out=io.StringIO();w=csv.writer(out);w.writerow(["receipt_code","arabic_name","english_name","rank","faculty","department","specialization","laboratory","submitted_at","publication_type","article_title","article_title_en","journal_name","classification","databases","publication_year","article_url","scholar_url","book_title","book_type","publisher","book_year","excerpt_url","cover_path"]);[w.writerow(list(r)) for r in rows]
    return send_file(io.BytesIO(out.getvalue().encode("utf-8-sig")),mimetype="text/csv",as_attachment=True,download_name="blida2_scientific_publications.csv")

@app.post("/logout")
@admin_required
def logout():
    check_csrf();session.clear();return redirect(url_for("admin"))

if __name__=="__main__":app.run(host="0.0.0.0",port=int(os.environ.get("PORT","5000")),debug=False)
