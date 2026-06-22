from __future__ import annotations

import cgi
import json
import mimetypes
import os
import shutil
import tempfile
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from processor import process_workbook


ROOT_DIR = Path(__file__).resolve().parent
STATIC_DIR = ROOT_DIR / "static"
HOST = "0.0.0.0"
PORT = int(os.environ.get("PORT", "8000"))
DOWNLOAD_FILENAME = "Maersk Rates Output.xlsx"


class AutomationHandler(BaseHTTPRequestHandler):
    server_version = "MaerskAutomation/1.0"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path

        if path in ("/", "/index.html"):
            self.serve_file(STATIC_DIR / "index.html", "text/html; charset=utf-8")
            return

        if path.startswith("/static/"):
            target = STATIC_DIR / path.removeprefix("/static/")
            self.serve_file(target)
            return

        self.send_error(HTTPStatus.NOT_FOUND, "File not found.")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path != "/api/process":
            self.send_error(HTTPStatus.NOT_FOUND, "Route not found.")
            return

        try:
            self.handle_process_request()
        except ValueError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except Exception as exc:  # pragma: no cover - safety net for local server use
            self.send_json({"error": f"Processing failed: {exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def handle_process_request(self) -> None:
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": self.headers.get("Content-Type", ""),
            },
        )

        file_field = form["file"] if "file" in form else None
        if file_field is None or not getattr(file_field, "filename", ""):
            raise ValueError("Please upload an Excel file first.")

        filename = Path(file_field.filename).name
        if Path(filename).suffix.lower() != ".xlsx":
            raise ValueError("Only .xlsx files are supported right now.")

        with tempfile.TemporaryDirectory(prefix="maersk_automation_") as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / filename
            output_path = temp_path / DOWNLOAD_FILENAME

            with input_path.open("wb") as output_stream:
                shutil.copyfileobj(file_field.file, output_stream)

            processed_file = process_workbook(input_path, output_path)
            payload = processed_file.read_bytes()

        download_name = DOWNLOAD_FILENAME
        self.send_response(HTTPStatus.OK)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.send_header("Content-Disposition", f'attachment; filename="{download_name}"')
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(payload)

    def serve_file(self, file_path: Path, content_type: str | None = None) -> None:
        if not file_path.exists() or not file_path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND, "File not found.")
            return

        mime_type = content_type or mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
        data = file_path.read_bytes()

        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", mime_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_json(self, payload: dict[str, str], status: HTTPStatus) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def run_server() -> None:
    with ThreadingHTTPServer((HOST, PORT), AutomationHandler) as server:
        print(f"Server running at http://{HOST}:{PORT}")
        server.serve_forever()


if __name__ == "__main__":
    run_server()
