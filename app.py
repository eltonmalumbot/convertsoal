import io
from pathlib import Path
from flask import Flask, request, jsonify, render_template, send_file, send_from_directory
from werkzeug.exceptions import RequestEntityTooLarge
from converter import parse_docx, parse_moodle_xml, parse_csv, export_data, ConversionError

# One WSGI app for Vercel Python Functions and local Flask development.
# Keep uploads/responses below Vercel Functions' 4.5 MB payload limit.
MAX_UPLOAD_BYTES = 3_800_000
MAX_PAYLOAD_BYTES = 4_000_000
app = Flask(__name__, static_folder=None)
app.config['MAX_CONTENT_LENGTH'] = MAX_PAYLOAD_BYTES

@app.after_request
def protect_exam_content(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['Referrer-Policy'] = 'no-referrer'
    if request.path.startswith('/api/'):
        response.headers['Cache-Control'] = 'no-store'
    return response

@app.get('/static/<path:filename>')
def local_static(filename):
    # Vercel serves public/static/* from its CDN before invoking Flask.
    # Local Flask keeps the same URL structure.
    return send_from_directory(Path(__file__).resolve().parent / 'public' / 'static', filename)


@app.get('/')
def index():
    return render_template('index.html')

@app.get('/api/health')
def health():
    return jsonify(status='ok', version='0.2.0-vercel', max_upload_bytes=MAX_UPLOAD_BYTES)

@app.post('/api/parse')
def parse():
    f = request.files.get('file')
    if not f or not f.filename:
        return jsonify(error='Pilih berkas DOCX, Moodle XML, atau CSV.'), 400
    name = f.filename.rsplit('/', 1)[-1].rsplit('\\', 1)[-1]
    suffix = name.lower().rsplit('.', 1)[-1]
    data = f.read()
    if len(data) > MAX_UPLOAD_BYTES:
        return jsonify(error='Maksimal 3,8 MB per berkas pada Vercel. Untuk file lebih besar perlu alur upload berbeda.'), 413
    try:
        if suffix == 'docx':
            result = parse_docx(data, name)
        elif suffix == 'xml':
            result = parse_moodle_xml(data, name)
        elif suffix == 'csv':
            result = parse_csv(data, name)
        else:
            return jsonify(error='Format masukan yang didukung: .docx, .xml, .csv.'), 400
    except (ConversionError, ValueError, KeyError) as exc:
        return jsonify(error=str(exc)), 422
    except Exception:
        app.logger.exception('Document parsing error')
        return jsonify(error='Berkas tidak dapat dibaca. Periksa format atau kerusakannya.'), 422
    # Large images can expand significantly when embedded as Base64 in JSON.
    # Check response too; otherwise Vercel returns an opaque 500/413 error.
    response = jsonify(result)
    if len(response.get_data()) > MAX_PAYLOAD_BYTES:
        return jsonify(error='Hasil pembacaan melebihi batas respons Vercel (4 MB). Kurangi ukuran gambar atau pecah dokumen.'), 413
    return response

@app.post('/api/export')
def export():
    body = request.get_json(silent=True)
    if not isinstance(body, dict) or not isinstance(body.get('questions'), list):
        return jsonify(error='Data soal tidak valid.'), 400
    if len(body['questions']) > 1500:
        return jsonify(error='Maksimal 1.500 soal per ekspor.'), 413
    fmt = str(body.get('format', 'xml')).lower()
    try:
        content, mimetype, ext = export_data(body['questions'], fmt, str(body.get('category', 'Bank Soal')))
    except ConversionError as exc:
        return jsonify(error=str(exc)), 422
    if len(content) > MAX_PAYLOAD_BYTES:
        return jsonify(error='Hasil ekspor melebihi batas Vercel (4 MB). Pecah bank soal menjadi beberapa file.'), 413
    return send_file(io.BytesIO(content), mimetype=mimetype,
                     as_attachment=True, download_name='moodle-question-studio.' + ext)

@app.errorhandler(RequestEntityTooLarge)
def too_large(_):
    return jsonify(error='Request terlalu besar untuk Vercel (batas aman 4 MB). Pecah dokumen atau soal.'), 413

if __name__ == '__main__':
    app.run(debug=False, port=8080, host='127.0.0.1')
