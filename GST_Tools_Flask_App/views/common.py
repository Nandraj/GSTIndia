from flask import (
    Blueprint,
    send_from_directory
)

from config import DOWNLOAD_FOLDER

common_bp = Blueprint('common_bp', __name__)


@common_bp.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(DOWNLOAD_FOLDER, filename, as_attachment=True)
