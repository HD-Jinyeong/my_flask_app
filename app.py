from flask import Flask, request

app = Flask(__name__)

# 홈 (기존)
@app.route("/")
def home():
    return "Hello, World! Render에서 잘 실행됩니다 🎉"

# 새로운 라우트: /upload
@app.route("/upload", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        f = request.files["file"]       # 업로드된 파일 가져오기
        f.save(f.filename)              # 서버에 파일 저장
        return f"'{f.filename}' 업로드 완료!"
    return """
        <h1>파일 업로드</h1>
        <form method="post" enctype="multipart/form-data">
            <input type="file" name="file">
            <input type="submit" value="업로드">
        </form>
    """

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

