from flask import Flask, render_template, request, send_file
from processor import process_files
from werkzeug.utils import secure_filename
import uuid

app = Flask(__name__)

# Maximum upload size (50MB)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

# Allowed file extensions
ALLOWED_EXTENSIONS = {"xlsx"}


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.errorhandler(413)
def file_too_large(e):
    return "File too large. Maximum upload size is 50MB.", 413


@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        try:
            files = request.files.getlist("files")

            if not files:
                return "No files uploaded"

            if len(files) > 20:
                return "Maximum 20 files allowed per upload"

            validated_files = []

            for file in files:

                if file.filename == "":
                    continue

                if not allowed_file(file.filename):
                    return f"Invalid file type: {file.filename}. Only .xlsx files are allowed."

                # Secure filename
                filename = secure_filename(file.filename)

                # Prevent duplicate filenames
                unique_name = f"{uuid.uuid4().hex}_{filename}"
                file.filename = unique_name

                validated_files.append(file)

            if not validated_files:
                return "No valid Excel files uploaded."

            output = process_files(validated_files)

            return send_file(
                output,
                as_attachment=True,
                download_name="transformed.xlsx",
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            return f"Error processing files: {str(e)}"

    return render_template("index.html")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)