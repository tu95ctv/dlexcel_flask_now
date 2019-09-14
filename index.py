# -*- coding: utf-8 -*-
from flask import Flask, Response
app = Flask(__name__)
import xlwt
from ne_nep import ne_nep_report_xl
# from diem_danh import diem_danh_report_xl
# from nn_chi_tiet import nn_chi_tiet_report_xl
from call_func import get_funcxl_and_run_funcxl_from_key
# try:
#     from StringIO import StringIO ## for Python 2
# except ImportError:
#     from io import StringIO ## for Python 3
    
from io import BytesIO
import mimetypes
from flask import Response
from werkzeug.datastructures import Headers
from datetime import datetime
from flask import request


@app.route('/')
@app.route('/index')
def index():
    return "Hello, World tu day"

@app.route('/age')
def age():
    from_ = request.args.get('from','not_from')
    return "Hello, World tu day: from %s"%from_

def get_func_kargs_():
    from_ = request.args.get('from','')
    to_ = request.args.get('to','')
    ne_nep_variable_values = {}
    if from_:
        ne_nep_variable_values['from'] = from_
    if to_:
        ne_nep_variable_values['to'] = to_
        
    kargs = {}
    kargs['variable_values'] = ne_nep_variable_values
    if 'font_size' in request.args:
        kargs['font_size'] = int(request.args['font_size'])
    if 'font' in request.args:
        font = request.args['font']
        kargs['font'] = font
        
        
    return kargs

    
#http://127.0.0.1:5000/dlxl/ne_nep?from=1987-09-22&to=2020-09-24
#http://127.0.0.1:5000/dlxl/diem_danh?from=1987-09-22&to=2020-09-24
#http://127.0.0.1:5000/dlxl/nn_chi_tiet?from=1987-09-22&to=2020-09-24
#http://127.0.0.1:5000/dlxl/nn_chi_tiet?from=1987-09-22&to=2020-09-24&font_size=12&font=2&break_sheet=false
@app.route('/dlxl/<func_key>')
def dlhaha(func_key):
#     func_key= request.args.get('func',None)
    if func_key == None:
        raise ValueError(u'không có tên hàm download xl')
    else:
        request_args = request.args
        wb, filename = get_funcxl_and_run_funcxl_from_key(func_key, request_args)
    
    response = Response()
    response.status_code = 200
    
    output = BytesIO() 
    wb.save(output)
    response.data = output.getvalue()
    mimetype_tuple = mimetypes.guess_type(filename)
    response_headers = Headers({
            'Pragma': "public",  # required,
            'Expires': '0',
            'Cache-Control': 'must-revalidate, post-check=0, pre-check=0',
            'Cache-Control': 'private',  # required for certain browsers,
            'Content-Type': mimetype_tuple[0],
            'Content-Disposition': 'attachment; filename=\"%s\";' % filename,
            'Content-Transfer-Encoding': 'binary',
            'Content-Length': len(response.data)
        })

    if not mimetype_tuple[1] is None:
        response.update({
                'Content-Encoding': mimetype_tuple[1]
            })

    response.headers = response_headers
    return response








