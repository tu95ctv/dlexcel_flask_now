# -*- coding: utf-8 -*-
from flask import Flask, Response
app = Flask(__name__)
import xlwt
from ne_nep import ne_nep_report_xl
from diem_danh import diem_danh_report_xl
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

dlxl_map_func = {'ne_nep':{'func':ne_nep_report_xl,'file_name':'ne_nep_report_xl','get_func_kargs':get_func_kargs_},
                 'diem_danh':{'func':diem_danh_report_xl,'file_name':'diem_danh_report','get_func_kargs':get_func_kargs_}
                 }

    
#http://127.0.0.1:5000/dlxl/ne_nep?from=1987-09-22&to=2020-09-24&font=1&font_size=13
#http://127.0.0.1:5000/dlxl/diem_danh?from=1987-09-22&to=2020-09-24&font_size=15
@app.route('/dlxl/<func_key>')
def dlhaha(func_key):
#     func_key= request.args.get('func',None)
    if func_key == None:
        raise ValueError(u'không có tên hàm download xl')
    else:
        adict = dlxl_map_func[func_key]
        func = adict['func']
        filename = adict['file_name'] +'.xls'
        get_func_kargs = adict.get('get_func_kargs',None)
    
    
    
    if get_func_kargs:
        kargs = get_func_kargs()
    else:
        kargs = None
    wb = func(**kargs)
    
    
    response = Response()
    response.status_code = 200

    output = BytesIO() 
#     wb = func()
    wb.save(output)
    response.data = output.getvalue()

#     filename = 'ne_nep.xls'
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
















# @app.route('/dlexcel')
# @app.route('/dlxl')
# def dlhaha():
#     from_ = request.args.get('from','')
#     to_ = request.args.get('to','')
#     if from_:
#         try:
#             from_ = datetime.strptime(from_,'%Y-%m-%d')
#             from_ = from_.strftime('%d/%m/%Y')
#         except ValueError:
#             from_ = ''
#             
#     if to_:
#         try:
#             to_ = datetime.strptime(to_,'%Y-%m-%d')
#             to_ = to_.strftime('%d/%m/%Y')
#         except ValueError:
#             to_ = ''
#             
#             
#         
#     #########################
#     # Code for creating Flask
#     # response
#     #########################
#     response = Response()
#     response.status_code = 200
# 
# 
#     ##################################
#     # Code for creating Excel data and
#     # inserting into Flask response
#     ##################################
# #     workbook = xlwt.Workbook()
#    
#     
#     style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
#     num_format_str='#,##0.00')
#     style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
#     
#     wb = xlwt.Workbook()
#     ws = wb.add_sheet('A Test Sheet')
#     
#     ws.write(0, 0, u'Từ ngày: %s đến %s'%(from_,to_), style0)
#     ws.write(1, 0, datetime.now(), style1)
#     ws.write(2, 0, 1)
#     ws.write(2, 1, 1)
#     ws.write(2, 2, xlwt.Formula("A3+B3"))
#     
# 
# 
#     #.... code here for adding worksheets and cells
# 
#     output = BytesIO() 
#     wb.save(output)
#     response.data = output.getvalue()
# 
#     ################################
#     # Code for setting correct
#     # headers for jquery.fileDownload
#     #################################
#     filename = 'export.xls'
#     mimetype_tuple = mimetypes.guess_type(filename)
# 
#     #HTTP headers for forcing file download
#     response_headers = Headers({
#             'Pragma': "public",  # required,
#             'Expires': '0',
#             'Cache-Control': 'must-revalidate, post-check=0, pre-check=0',
#             'Cache-Control': 'private',  # required for certain browsers,
#             'Content-Type': mimetype_tuple[0],
#             'Content-Disposition': 'attachment; filename=\"%s\";' % filename,
#             'Content-Transfer-Encoding': 'binary',
#             'Content-Length': len(response.data)
#         })
# 
#     if not mimetype_tuple[1] is None:
#         response.update({
#                 'Content-Encoding': mimetype_tuple[1]
#             })
# 
#     response.headers = response_headers
# 
#     #as per jquery.fileDownload.js requirements
# #     response.set_cookie('fileDownload', 'true', path='/')
# 
#     ################################
#     # Return the response
#     #################################
#     #return  send_file('/var/www/PythonProgramming/PythonProgramming/static/images/python.jpg', attachment_filename='python.jpg')
#     return response
    

