"""Optional local wire test for TLS, static assets and the unconfigured API."""
import json
import os
from pathlib import Path
import ssl
import subprocess
import tempfile
from threading import Thread
import unittest
from unittest.mock import patch
from urllib.error import HTTPError
from urllib.request import Request, urlopen

from server import Handler, ThreadingHTTPServer


@unittest.skipUnless(os.getenv("SHEETGUARD_TEST_HTTPS") == "1", "Requires loopback sockets and OpenSSL")
class HttpsTests(unittest.TestCase):
    def test_local_tls_service(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(os.environ, {"SHEETGUARD_MODE":"development"}, clear=True):
            root = Path(tmp)
            cert, key = root / "dev.pem", root / "dev.key"
            subprocess.run(["/usr/bin/openssl", "req", "-new", "-newkey", "rsa:2048", "-x509", "-nodes",
                            "-days", "1", "-subj", "/CN=localhost", "-keyout", str(key), "-out", str(cert)],
                           check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            with ThreadingHTTPServer(("127.0.0.1", 0), Handler) as service:
                context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
                context.load_cert_chain(cert, key)
                service.socket = context.wrap_socket(service.socket, server_side=True)
                worker = Thread(target=service.serve_forever, daemon=True)
                worker.start()
                try:
                    origin = f"https://localhost:{service.server_port}"
                    trust = ssl.create_default_context(cafile=str(cert))
                    for resource, media in [("/health","application/json"),("/index.html","text/html"),
                                            ("/app.mjs","text/javascript"),("/icon-32.png","image/png")]:
                        with urlopen(origin+resource,context=trust,timeout=5) as response:
                            self.assertEqual(response.headers.get_content_type(), media)
                            self.assertTrue(response.read())
                    req=Request(origin+"/api/agent",data=b'{"question":"Audit my data"}',headers={"Content-Type":"application/json"})
                    with self.assertRaises(HTTPError) as error:
                        urlopen(req,context=trust,timeout=5)
                    with error.exception as response:
                        self.assertEqual(response.code,503)
                        self.assertIn("unconfigured",json.loads(response.read())["error"])
                finally:
                    service.shutdown()
                    worker.join()
