# -*- coding: utf-8 -*-
from flask import Flask, Response
app = Flask(__name__)
import xlwt

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



@app.route('/dlexcel')
@app.route('/dlxl')
def dlhaha():
    from_ = request.args.get('from','')
    to_ = request.args.get('to','')
    #########################
    # Code for creating Flask
    # response
    #########################
    response = Response()
    response.status_code = 200


    ##################################
    # Code for creating Excel data and
    # inserting into Flask response
    ##################################
#     workbook = xlwt.Workbook()
   
    
    style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
    num_format_str='#,##0.00')
    style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
    
    wb = xlwt.Workbook()
    ws = wb.add_sheet('A Test Sheet')
    
    ws.write(0, 0, u'Từ ngày: %s đến %s'%(from_,to_), style0)
    ws.write(1, 0, datetime.now(), style1)
    ws.write(2, 0, 1)
    ws.write(2, 1, 1)
    ws.write(2, 2, xlwt.Formula("A3+B3"))
    


    #.... code here for adding worksheets and cells

    output = BytesIO() 
    wb.save(output)
    response.data = output.getvalue()

    ################################
    # Code for setting correct
    # headers for jquery.fileDownload
    #################################
    filename = 'export.xls'
    mimetype_tuple = mimetypes.guess_type(filename)

    #HTTP headers for forcing file download
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

    #as per jquery.fileDownload.js requirements
#     response.set_cookie('fileDownload', 'true', path='/')

    ################################
    # Return the response
    #################################
    #return  send_file('/var/www/PythonProgramming/PythonProgramming/static/images/python.jpg', attachment_filename='python.jpg')
    return response
    

