import http.server, ssl


class RequestHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory='./FractalExcelWebAddInWeb', **kwargs)


server_address = ('localhost', 8080)
httpd = http.server.HTTPServer(server_address, RequestHandler)
httpd.socket = ssl.wrap_socket(httpd.socket,
                               server_side=True,
                               certfile='localhost_cert\localhost_cert.pem',
                               keyfile='localhost_cert\localhost_key.pem',
                               ssl_version=ssl.PROTOCOL_TLS)
print('Started server on https://127.0.0.1:8080')
httpd.serve_forever()