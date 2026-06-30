"""Flask Web アプリ — 旅行命令復命書 Excel をアップロードして動静表 Excel をダウンロード."""

from __future__ import annotations

import os
from io import BytesIO

from flask import Flask, abort, render_template, request, send_file

import processor

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16MB


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/healthz")
def healthz():
    return "ok", 200


@app.post("/process")
def process_route():
    file = request.files.get("file")
    if file is None or not file.filename:
        abort(400, "ファイルが選択されていません")
    if not file.filename.lower().endswith(".xlsx"):
        abort(400, ".xlsx ファイルを指定してください")

    order_raw = request.form.get("order", "")
    order = [k.strip() for k in order_raw.split(",") if k.strip()] or None

    try:
        out_bytes, out_name, _days, _entries = processor.process_excel(file.stream, order=order)
    except ValueError as e:
        abort(400, str(e))
    except Exception as e:  # noqa: BLE001
        abort(500, f"処理に失敗しました: {e}")

    return send_file(
        BytesIO(out_bytes),
        as_attachment=True,
        download_name=out_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
