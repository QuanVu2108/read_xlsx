import json
from flask import Flask, Response, abort
from .read_xlsx import get_data, JSON_MIME_TYPE

app = Flask(__name__)

data = dict()

# @app.route('/all')
# def book_list():
	# data = get_data(all)
	# response = Response(json.dumps(data, indent=4, sort_keys=True), status=200, mimetype=JSON_MIME_TYPE)
	# return response


@app.route('/<string:col_id>')
def data_detail(col_id):
    if col_id != 'all' and len(col_id) > 2:
        abort(404)
    data = get_data(col_id)
    if data is None:
        abort(404)

    content = json.dumps(data, indent=4, sort_keys=False)
    return content, 200, {'Content-Type': JSON_MIME_TYPE}


@app.errorhandler(404)
def not_found(e):
    return '', 404
