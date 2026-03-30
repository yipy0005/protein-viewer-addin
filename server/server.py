"""
Protein Viewer — Lightweight static file server.
Serves the webpack dist/ output for the PowerPoint add-in.
Uses a persistent self-signed HTTPS certificate stored in ~/.protein-viewer/certs/
and adds it to the macOS Keychain as trusted on first run.
"""

import os
import sys
import subprocess
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request

app = FastAPI(title="Protein Viewer Server", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


class NoCacheMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        response = await call_next(request)
        if request.url.path.endswith((".js", ".html", ".json", ".css")):
            response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate"
            response.headers["Pragma"] = "no-cache"
        return response


app.add_middleware(NoCacheMiddleware)


@app.get("/health")
def health_check():
    return {"status": "ok"}


def get_dist_path():
    if getattr(sys, "frozen", False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "dist")


dist_path = get_dist_path()
if os.path.isdir(dist_path):
    app.mount("/", StaticFiles(directory=dist_path, html=True), name="static")
else:
    print(f"WARNING: dist/ directory not found at {dist_path}")


def get_cert_dir():
    cert_dir = os.path.join(os.path.expanduser("~"), ".protein-viewer", "certs")
    os.makedirs(cert_dir, exist_ok=True)
    return cert_dir


def generate_self_signed_cert():
    cert_dir = get_cert_dir()
    cert_path = os.path.join(cert_dir, "localhost.crt")
    key_path = os.path.join(cert_dir, "localhost.key")

    if os.path.exists(cert_path) and os.path.exists(key_path):
        print(f"Using existing SSL certificate from {cert_dir}")
        return cert_path, key_path

    try:
        print("Generating self-signed SSL certificate for localhost...")
        subprocess.run(
            [
                "openssl", "req", "-x509", "-newkey", "rsa:2048",
                "-keyout", key_path, "-out", cert_path,
                "-days", "3650", "-nodes",
                "-subj", "/CN=localhost",
                "-addext", "subjectAltName=DNS:localhost,IP:127.0.0.1",
            ],
            check=True, capture_output=True, timeout=10,
        )
        if not (os.path.exists(cert_path) and os.path.exists(key_path)):
            return None, None
        trust_certificate_macos(cert_path)
        return cert_path, key_path
    except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired) as e:
        print(f"SSL cert generation failed: {e}")
        return None, None


def trust_certificate_macos(cert_path):
    import platform
    if platform.system() != "Darwin":
        return
    try:
        print("Adding certificate to macOS Keychain...")
        subprocess.run(
            ["security", "add-trusted-cert", "-r", "trustRoot",
             "-k", os.path.expanduser("~/Library/Keychains/login.keychain-db"),
             cert_path],
            check=True, capture_output=True, timeout=30,
        )
        print("Certificate trusted in Keychain.")
    except subprocess.CalledProcessError:
        try:
            subprocess.run(
                ["security", "add-trusted-cert", "-r", "trustRoot", cert_path],
                check=True, capture_output=True, timeout=30,
            )
            print("Certificate trusted in Keychain.")
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as e:
            print(f"Could not auto-trust certificate: {e}")
    except (FileNotFoundError, subprocess.TimeoutExpired) as e:
        print(f"Could not auto-trust certificate: {e}")
