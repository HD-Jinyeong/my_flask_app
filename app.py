from flask import Flask
app = Flask(__name__)

@app.route("/")
def home():
    return "Hello, World! Render에서 잘 실행됩니다 🎉"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
