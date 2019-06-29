import json
from flask import Flask, Response, abort
from .read_xlsx import get_data, JSON_MIME_TYPE

app = Flask(__name__)

data = dict()

@app.route('/<string:file_id>/<string:col_id>')
def data_detail(file_id, col_id):
    data = get_data(file_id, col_id)
    if data is None:
        abort(404)

    content = json.dumps(data, indent=4, sort_keys=False)
    return content, 200, {'Content-Type': JSON_MIME_TYPE}


@app.errorhandler(404)
def not_found(e):
    return '', 404