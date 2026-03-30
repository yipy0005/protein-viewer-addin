"""Entry point for the Protein Viewer server."""

import os
import sys


def main():
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from server import app, generate_self_signed_cert
    import uvicorn

    port = 3001

    cert_path, key_path = generate_self_signed_cert()

    if cert_path and key_path:
        print(f"Starting Protein Viewer server on https://localhost:{port}")
        uvicorn.run(app, host="127.0.0.1", port=port, log_level="warning",
                    ssl_certfile=cert_path, ssl_keyfile=key_path)
    else:
        print(f"Starting Protein Viewer server on http://localhost:{port} (no SSL)")
        uvicorn.run(app, host="127.0.0.1", port=port, log_level="warning")


if __name__ == "__main__":
    main()
