import os
from werkzeug.utils import secure_filename
from flask import current_app

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in {"png", "jpg", "jpeg", "gif"}

def save_profile_image(file, username):
    if file and allowed_file(file.filename):
        filename = secure_filename(f"{username}_{file.filename}")
        filepath = os.path.join(current_app.config["UPLOAD_FOLDER"], filename)
        os.makedirs(current_app.config["UPLOAD_FOLDER"], exist_ok=True)
        file.save(filepath)
        return filename
    return None