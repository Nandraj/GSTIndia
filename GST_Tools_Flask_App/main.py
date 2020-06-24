from flask import Flask
from views.general import general_views_bp
from views.refund import refund_tools_bp
from views.common import common_bp

app = Flask(__name__)

app.register_blueprint(general_views_bp)
app.register_blueprint(refund_tools_bp)
app.register_blueprint(common_bp)

if __name__ == "__main__":
    app.run(debug=True)
