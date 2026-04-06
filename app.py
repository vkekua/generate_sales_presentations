import io
import os
import socket
import shutil
import tempfile
import zipfile

from flask import Flask, request, send_file, render_template_string

from main import generate_all

app = Flask(__name__)

TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sales Presentation Generator</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: #1a1a1a;
            color: #fff;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .container {
            background: #242424;
            border-radius: 12px;
            padding: 48px;
            max-width: 480px;
            width: 100%;
            text-align: center;
            box-shadow: 0 4px 24px rgba(0,0,0,0.4);
        }
        h1 {
            font-size: 24px;
            margin-bottom: 8px;
            color: #f9d605;
        }
        .subtitle {
            color: #999;
            font-size: 14px;
            margin-bottom: 32px;
        }
        .upload-area {
            border: 2px dashed #444;
            border-radius: 8px;
            padding: 32px 16px;
            margin-bottom: 24px;
            transition: border-color 0.2s;
        }
        .upload-area:hover { border-color: #f9d605; }
        .upload-area label {
            cursor: pointer;
            color: #ccc;
            font-size: 14px;
        }
        .upload-area input[type="file"] { display: none; }
        .file-name {
            margin-top: 12px;
            font-size: 13px;
            color: #f9d605;
            min-height: 20px;
        }
        button {
            background: #f9d605;
            color: #1a1a1a;
            border: none;
            border-radius: 8px;
            padding: 14px 32px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            width: 100%;
            transition: opacity 0.2s;
        }
        button:hover { opacity: 0.9; }
        button:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }
        .error {
            background: #3a1a1a;
            border: 1px solid #ff6257;
            color: #ff6257;
            border-radius: 8px;
            padding: 12px;
            margin-bottom: 24px;
            font-size: 14px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Sales Presentation Generator</h1>
        <p class="subtitle">Upload your Excel file to generate presentations</p>

        {% if error %}
        <div class="error">{{ error }}</div>
        {% endif %}

        <form method="POST" action="/generate" enctype="multipart/form-data" id="genForm">
            <div class="upload-area">
                <label>
                    Click to select Excel file (.xlsx)
                    <input type="file" name="file" accept=".xlsx,.xls" id="fileInput">
                </label>
                <div class="file-name" id="fileName"></div>
            </div>
            <button type="submit" id="submitBtn">Generate Presentations</button>
        </form>
    </div>

    <script>
        const fileInput = document.getElementById('fileInput');
        const fileName = document.getElementById('fileName');
        const form = document.getElementById('genForm');
        const btn = document.getElementById('submitBtn');

        fileInput.addEventListener('change', () => {
            fileName.textContent = fileInput.files[0]?.name || '';
        });

        form.addEventListener('submit', (e) => {
            if (!fileInput.files.length) {
                e.preventDefault();
                fileName.textContent = 'Please select a file first.';
                return;
            }
            btn.disabled = true;
            btn.textContent = 'Generating... please wait';
        });
    </script>
</body>
</html>
"""


def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


@app.route("/")
def index():
    error = request.args.get("error")
    return render_template_string(TEMPLATE, error=error)


@app.route("/generate", methods=["POST"])
def generate():
    file = request.files.get("file")
    if not file or file.filename == "":
        return render_template_string(TEMPLATE, error="No file selected.")

    if not file.filename.lower().endswith((".xlsx", ".xls")):
        return render_template_string(TEMPLATE, error="Please upload an Excel file (.xlsx or .xls).")

    tmp_dir = tempfile.mkdtemp()
    try:
        input_path = os.path.join(tmp_dir, "input.xlsx")
        output_dir = os.path.join(tmp_dir, "output")
        file.save(input_path)

        generated = generate_all(input_path, output_dir)

        if not generated:
            return render_template_string(TEMPLATE, error="No presentations were generated. Check that your Excel file has partners with CreatePPT = True.")

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for path in generated:
                zf.write(path, os.path.basename(path))
        buf.seek(0)

        return send_file(buf, mimetype="application/zip", as_attachment=True, download_name="presentations.zip")
    except Exception as e:
        return render_template_string(TEMPLATE, error=f"Generation failed: {e}")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    local_ip = get_local_ip()
    port = 5000

    cert_file = os.path.join(os.path.dirname(__file__), "cert.pem")
    key_file = os.path.join(os.path.dirname(__file__), "key.pem")

    if os.path.exists(cert_file) and os.path.exists(key_file):
        protocol = "https"
        ssl_ctx = (cert_file, key_file)
    else:
        protocol = "http"
        ssl_ctx = None
        print("  WARNING: cert.pem/key.pem not found. Running without HTTPS.")
        print("  Run 'python generate_cert.py' to enable HTTPS.\n")

    print(f"\n  Sales Presentation Generator")
    print(f"  Local:   {protocol}://localhost:{port}")
    print(f"  Network: {protocol}://{local_ip}:{port}\n")
    app.run(host="0.0.0.0", port=port, ssl_context=ssl_ctx, threaded=True)
