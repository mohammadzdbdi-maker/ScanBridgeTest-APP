import http.server
import ssl

httpd = http.server.HTTPServer(('0.0.0.0', 8000), http.server.SimpleHTTPRequestHandler)
context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
context.load_cert_chain('cert.pem', 'key.pem')
httpd.socket = context.wrap_socket(httpd.socket, server_side=True)
print("سرور HTTPS روی پورت 8000 روشن شد")
httpd.serve_forever()