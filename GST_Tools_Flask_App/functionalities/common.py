# Common Purpose variables
ALLOWED_EXTENSIONS = ['xls', 'xlsx']


# Gen purpose functions
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
