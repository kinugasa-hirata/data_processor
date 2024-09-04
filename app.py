from flask import Flask, render_template, jsonify

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/hello')
def hello():
    return jsonify(message="Hello from the backend!")

if __name__ == '__main__':
    app.run(debug=True)
