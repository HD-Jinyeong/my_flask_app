import os
import boto3
from flask import Flask, request

app = Flask(__name__)

# Render 환경변수에서 불러오기
S3_BUCKET = os.getenv("S3_BUCKET")
S3_REGION = os.getenv("S3_REGION")
AWS_ACCESS_KEY_ID = os.getenv("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = os.getenv("AWS_SECRET_ACCESS_KEY")

# boto3 클라이언트 생성
s3 = boto3.client(
    "s3",
    aws_access_key_id=AWS_ACCESS_KEY_ID,
    aws_secret_access_key=AWS_SECRET_ACCESS_KEY,
    region_name=S3_REGION
)

@app.route("/", methods=["GET", "POST"])
def home():
    if request.method == "POST":
        f = request.files["file"]

        # 파일을 S3에 업로드
        s3.upload_fileobj(f, S3_BUCKET, f.filename)

        # 업로드된 파일 URL 생성
        file_url = f"https://{S3_BUCKET}.s3.{S3_REGION}.amazonaws.com/{f.filename}"
        return f"""
            <h3>S3 업로드 성공!</h3>
            <p>파일 이름: {f.filename}</p>
            <p>URL: <a href="{file_url}" target="_blank">{file_url}</a></p>
        """
    return """
        <h1>AWS S3 파일 업로드</h1>
        <form method="post" enctype="multipart/form-data">
            <input type="file" name="file"><br><br>
            <input type="submit" value="업로드">
        </form>
    """

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
