import json
import os
import uuid
import puremagic
from datetime import datetime, timedelta
from pathlib import Path

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.utils import secure_filename

BASE_DIR   = Path(__file__).parent
PLANTILLA  = BASE_DIR / 'resources' / 'plantilla.docx'
UPLOAD_DIR = BASE_DIR / 'uploads'
UPLOAD_DIR.mkdir(exist_ok=True)

from corrector import InDretCorrector

ALLOWED_EXT = {'.docx'}
ALLOWED_MIME = {
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/zip',  # DOCX és un contenidor ZIP; puremagic pot retornar aquest
}
MAX_MB      = 20

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['MAX_CONTENT_LENGTH'] = MAX_MB * 1024 * 1024


# ─── Neteja de fitxers antics (>30 min) ──────────────────────────────────────

def validar_mime(path: Path) -> bool:
    """Retorna True si el fitxer sembla un DOCX segons els bytes màgics."""
    try:
        matches = puremagic.magic_file(str(path))
        if not matches:
            return False
        return matches[0].mime_type in ALLOWED_MIME
    except puremagic.MagicException:
        return False


def netejar_uploads():
    limit = datetime.now() - timedelta(minutes=30)
    for f in UPLOAD_DIR.iterdir():
        if f.is_file() and f.name != '.gitkeep':
            mtime = datetime.fromtimestamp(f.stat().st_mtime)
            if mtime < limit:
                f.unlink(missing_ok=True)


# ─── Errors ──────────────────────────────────────────────────────────────────

@app.errorhandler(RequestEntityTooLarge)
def fitxer_massa_gran(e):
    flash(f'El fitxer supera els {MAX_MB} MB màxims permesos.')
    return redirect(url_for('index'))


# ─── Rutes ───────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    netejar_uploads()
    return render_template('index.html')


@app.route('/corregir', methods=['POST'])
def corregir():
    fitxer = request.files.get('fitxer')

    if not fitxer or fitxer.filename == '':
        flash('Selecciona un fitxer .docx abans de continuar.')
        return redirect(url_for('index'))

    ext = Path(fitxer.filename).suffix.lower()
    if ext not in ALLOWED_EXT:
        flash('Només s\'accepten fitxers .docx.')
        return redirect(url_for('index'))

    sid        = uuid.uuid4().hex
    nom_segur  = secure_filename(fitxer.filename)
    input_path = UPLOAD_DIR / f"{sid}_input_{nom_segur}"
    fitxer.save(input_path)

    if not validar_mime(input_path):
        input_path.unlink(missing_ok=True)
        flash('El fitxer no és un document .docx vàlid.')
        return redirect(url_for('index'))

    edicio       = request.form.get('edicio', '').strip()
    doi          = request.form.get('doi', '').strip()
    recepcio     = request.form.get('recepcio', '').strip()
    acceptacio   = request.form.get('acceptacio', '').strip()
    try:
        pagina_inici = int(request.form.get('pagina_inici', '1') or '1')
        if pagina_inici < 1:
            pagina_inici = 1
    except ValueError:
        pagina_inici = 1

    # Autors: llistes paral·leles autors_nom[] i autors_org[]
    noms = request.form.getlist('autors_nom[]')
    orgs = request.form.getlist('autors_org[]')
    autors = [
        {'nom': n.strip(), 'org': orgs[i].strip() if i < len(orgs) else ''}
        for i, n in enumerate(noms) if n.strip()
    ]
    if not autors:
        autor_legacy = request.form.get('autor', '').strip()
        if autor_legacy:
            autors = [{'nom': autor_legacy, 'org': ''}]

    stem = Path(nom_segur).stem

    try:
        corrector = InDretCorrector(str(input_path))
        if PLANTILLA.exists():
            out_doc, out_report = corrector.template_run(
                str(PLANTILLA), edicio, doi,
                recepcio=recepcio, acceptacio=acceptacio,
                pagina_inici=pagina_inici,
                autors=autors if autors else None,
            )
        else:
            out_doc, out_report = corrector.run()
    except Exception as exc:
        input_path.unlink(missing_ok=True)
        flash(f'Error en processar el document: {exc}')
        return redirect(url_for('index'))

    # Mou el .docx corregit a un nom controlat per sid
    out_doc_path = UPLOAD_DIR / f"{sid}_corregit.docx"
    Path(out_doc).rename(out_doc_path)

    # Esborra fitxers auxiliars generats pel corrector
    input_path.unlink(missing_ok=True)
    Path(out_report).unlink(missing_ok=True)

    # Desa l'informe com a JSON
    report_data = {
        'nom':      nom_segur,
        'stem':     stem,
        'applied':  corrector.report.applied,
        'alerts':   corrector.report.alerts,
        'pending':  corrector.report.pending,
        'edicio':     edicio,
        'doi':        doi,
        'autors':     autors,
        'recepcio':   recepcio,
        'acceptacio': acceptacio,
        'data':     datetime.now().strftime('%d/%m/%Y %H:%M'),
    }
    report_json = UPLOAD_DIR / f"{sid}_report.json"
    report_json.write_text(json.dumps(report_data, ensure_ascii=False), encoding='utf-8')

    return redirect(url_for('resultat', sid=sid))


@app.route('/resultat/<sid>')
def resultat(sid):
    report_json = UPLOAD_DIR / f"{sid}_report.json"
    out_doc     = UPLOAD_DIR / f"{sid}_corregit.docx"

    if not report_json.exists() or not out_doc.exists():
        flash('La sessió ha expirat o no existeix. Torna a pujar el document.')
        return redirect(url_for('index'))

    data = json.loads(report_json.read_text(encoding='utf-8'))
    return render_template('resultat.html', sid=sid, **data)


@app.route('/descarregar/<sid>')
def descarregar(sid):
    out_doc = UPLOAD_DIR / f"{sid}_corregit.docx"
    if not out_doc.exists():
        flash('El fitxer ja no està disponible. Torna a processar el document.')
        return redirect(url_for('index'))

    report_json = UPLOAD_DIR / f"{sid}_report.json"
    stem = 'article'
    if report_json.exists():
        data = json.loads(report_json.read_text(encoding='utf-8'))
        stem = data.get('stem', 'article')

    return send_file(
        out_doc,
        as_attachment=True,
        download_name=f"{stem}_corregit.docx",
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )


@app.route('/informe/<sid>')
def descarregar_informe(sid):
    report_json = UPLOAD_DIR / f"{sid}_report.json"
    if not report_json.exists():
        flash('L\'informe ja no està disponible.')
        return redirect(url_for('index'))

    data  = json.loads(report_json.read_text(encoding='utf-8'))
    lines = [
        f"Informe de correcció — {data['nom']} — {data['data']}",
        "=" * 60,
        "",
        "CANVIS APLICATS AUTOMÀTICAMENT",
        "-" * 40,
    ]
    for item in data['applied']:
        lines.append(f"[✓] {item}")
    lines += ["", "ALERTES QUE REQUEREIXEN REVISIÓ MANUAL", "-" * 40]
    if data['alerts']:
        for item in data['alerts']:
            lines.append(f"[!] {item}")
    else:
        lines.append("Cap alerta detectada.")
    lines += ["", "ELEMENTS INCOMPLETS (AFEGIR PER INDRET)", "-" * 40]
    for item in data['pending']:
        lines.append(f"[ ] {item}")

    txt = "\n".join(lines)
    stem = data.get('stem', 'article')

    from flask import Response
    return Response(
        txt,
        mimetype='text/plain; charset=utf-8',
        headers={'Content-Disposition': f'attachment; filename="{stem}_informe.txt"'}
    )


if __name__ == '__main__':
    app.run(debug=True, port=5000)
