from flask import Flask, render_template
from flask_socketio import SocketIO, emit, Namespace

class Server(Namespace):
    def __init__(self, host='localhost', port=5000, namespace="/"):
        self.namespace = namespace
        super().__init__(namespace)
        self.app = Flask(__name__)
        self.socketio = SocketIO(self.app)
        self.host = host
        self.port = port

        # self.configure_routes()

    # def configure_routes(self):
    #     # @self.app.route('/')
    #     # def index():
    #     #     return render_template('index.html')

    #     @self.socketio.on('message')
    #     def handle_message(data):
    #         print('received message: ' + data)
    #         # emit('response', {'data': 'Message received!'})
    #         emit('response', "message received")

    def on_message(self, data):
        print('Received message:' + data)
        emit('response', "message received")

    def run(self):
        self.socketio.on_namespace(self)
        self.socketio.run(self.app, host=self.host, port=self.port)

if __name__ == '__main__':
    server = Server()
    server.run()