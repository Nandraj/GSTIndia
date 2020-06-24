from flask import (
    Blueprint,
    render_template,
    request,
    redirect,
    url_for
)
from functionalities.common import (
    allowed_file
)
from werkzeug.utils import secure_filename
from functionalities.GSTR2_XL_Generator import gstr2_xl_generator
from functionalities.Stmt_1A_Json_Generator import stmt_1a_json_generator
from config import UPLOAD_FOLDER, DOWNLOAD_FOLDER
import sys
import os
sys.path.append(os.path.abspath(os.path.join('..', 'config')))

refund_tools_bp = Blueprint("refund_tools_bp", __name__)


@refund_tools_bp.route('/gstr2a_xl', methods=['GET', 'POST'])
def gstr2xl_upload():
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template('refund/gstr2a_xl.html', error="No file attached in request")
        file = request.files['file']
        if file.filename == '':
            return render_template('refund/gstr2a_xl.html', error="No file selected")
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(UPLOAD_FOLDER, filename))
            gstr2_xl_generator(os.path.join(
                UPLOAD_FOLDER, filename), DOWNLOAD_FOLDER)
            return redirect(url_for('common_bp.download_file', filename=filename))
        else:
            # file not in allowed file
            return render_template('refund/gstr2a_xl.html', error="Select valid excel file")
    return render_template('refund/gstr2a_xl.html')


@refund_tools_bp.route('/stmt1a_json', methods=['GET', 'POST'])
def stmt1a_json_upload():
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template('refund/stmt1a_json.html', error="No file attached in request")
        file = request.files['file']
        if file.filename == '':
            return render_template('refund/stmt1a_json.html', error="No file selected")
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            try:
                json_file_name = secure_filename(
                    (file.filename).replace(".xlsx", ".json"))
            except:
                json_file_name = secure_filename(
                    (file.filename).replace(".xls", ".json"))

            file.save(os.path.join(UPLOAD_FOLDER, filename))
            stmt_1a_json_generator(os.path.join(
                UPLOAD_FOLDER, filename), DOWNLOAD_FOLDER)
            return redirect(url_for('common_bp.download_file', filename=json_file_name))
        else:
            # file not in allowed file
            return render_template('refund/stmt1a_json.html', error="Select valid excel file")
    return render_template('refund/stmt1a_json.html')
