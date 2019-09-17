# -*- coding: utf-8 -*-
from flask import Flask, Response
from operator import itemgetter
app = Flask(__name__)
from io import BytesIO
import mimetypes
from werkzeug.datastructures import Headers
from flask import request

########################

import requests
from collections import  OrderedDict
from time import sleep
import xlwt
from datetime import datetime
import re
def font_decorator_parent_new(**kkgargs):
    def font_decorator(func):
        def awrapper(*args, **kargs):
            for k,v  in kkgargs.items():
                if k not in kargs and v:
                    kargs[k] = v
            rs = func(*args, **kargs)
            return rs
        return awrapper
    return font_decorator
def generate_easyxf_import (font='Times New Roman', 
                     bold = False,
                     underline=False,
                     height=12, 
                     align_wrap = False,
                     vert = False,
                     horiz = False,
                     borders = False,
                     pattern = False,
                     italic= False,**kargs
                     ):
    fonts = []
    fonts.append('name %s'%font)
    if underline:
        fonts.append('underline on')
    if bold:
        fonts.append('bold on')
        
    if italic:
        fonts.append('italic on')
        
    fonts.append('height %s'%(height*20))
    sums = []
    font = 'font: ' + ','.join(fonts)
    sums.append(font)
    
    aligns = []
    if vert:
        aligns.append('vert %s'%vert)
    if horiz:
        aligns.append('horiz %s'%horiz)
    if align_wrap:
        aligns.append('wrap on')
        
    if aligns:
        align = 'align:  ' + ','.join(aligns)
#         font = font + '; ' + align
        sums.append(align)
    
  
    if borders:
        borders = 'borders: ' + borders
        sums.append(borders) 
    
    if pattern:
        pattern = 'pattern: ' + pattern
        sums.append(pattern)
#     for k,v in kargs.items():
#         sums.append(k+': ' + v)
    sums = ';'.join(sums)   
    return sums


def get_hasura_data(data):
    url = 'https://qlth.hpz.vn/v1/graphql'
    headers = {'x-hasura-admin-secret': 'hpz', 'content-type': 'application/json', 'User-Agent' : 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.101 Safari/537.36'}
    count_fail = 0
    while 1:
        print ('get html',url)
        try:
            request = requests.post(url, json=data, headers=headers)
            return request.json()
        except Exception as e:
            count_fail +=1
            print ('loi khi get html',e)
            sleep(5)
            if count_fail ==5:
                raise ValueError(u'Lỗi get html')

def write_fixups(fixups_setting):
    fixups = fixups_setting['fixups']
    defaut_fixups_style = fixups_setting['default_fixups_style']
    wb = xlwt.Workbook()
    sheet_name = fixups_setting.get('sheet_name','Sheet 1')
    ws = wb.add_sheet(sheet_name, cell_overwrite_ok=True)
    irow = 0
    fixups = OrderedDict (fixups)
    height = fixups_setting.get('row_height')
    for k_fixups, v_fixups in fixups.items():
        skip_row = v_fixups.get('skip_row')
        if skip_row:
            continue
        break_sheet = v_fixups.get('break_sheet')
        if break_sheet:
            sheet_name = v_fixups.get('sheet_name','Sheet 1')
            ws = wb.add_sheet(sheet_name, cell_overwrite_ok=True)
            irow = 0
        row_height = v_fixups.get('row_height')
        height = row_height or height
        xrange = v_fixups['range']
        if xrange[0] == 'auto':
            offset = v_fixups.get('offset', 1 )
            irow = irow + offset
            xrange[0] = irow
            if len(xrange) == 4:
                irow = irow + xrange[1]
                xrange[1] = irow
        else:
            irow = xrange[0] 
        func = v_fixups.get('func')
        if func:
            begin_icol = xrange[1]
            func_kargs = v_fixups.get('func_kargs',{})
            func_row = func(ws, irow, begin_icol,**func_kargs)
            if func_row:
                irow = irow + func_row -1
        else:
            val = v_fixups.get('val',None)
            val_func = v_fixups.get('val_func')
            if val_func:
                val_func_kargs = v_fixups.get('val_func_kargs',{})
                val = val_func(**val_func_kargs)
                
            style = v_fixups.get('style',defaut_fixups_style)
            if len(xrange) == 2:
                ws.write(xrange[0], xrange[1], val, style)
            else:
                ws.write_merge(xrange[0], xrange[1],xrange[2], xrange[3], val, style)
                
            if height != None:
                ws.row(irow).height_mismatch = True
                ws.row(irow).height = height
    return wb
 
def get_hasura_data_with_query_and_variable(variable_values=None, query=None):
    data = {'query': query}
    if variable_values:
        data['variables'] = variable_values
    rs = get_hasura_data(data)
    return  rs
Convert_number_to_APB = {0:'A', 1:'B', 2:'C', 3:'D', 4:'E', 5:'F',6:'G',7:'H',8:'I',9:'J', 10:'K'}

# def write_table(table_setting, datas, begin_title_irow,  begin_icol, ws,gen_row_data = None ):
def write_table_rerange(ws, begin_title_irow, begin_icol, table_setting=None):
    title_height = table_setting.get('title_height')
    datas = table_setting['datas']
    def get_width(num_characters,font_height=12):
        return int((1+num_characters) * 256*font_height/12)
    height = table_setting.get('row_height')
    default_cell_font = table_setting['default_cell_font']
    skip_width = table_setting.get('skip_width')
    all_cell_func = table_setting.get('all_cell_func')
    def write_a_row(obj_data, FIELDNAME_FIELDATTR,  ws, irow, begin_icol,default_cell_font=default_cell_font, write_a_row_more_data = {}):
        
        ifield = 0 
        for fname, field_attr_dict  in FIELDNAME_FIELDATTR.items():
            icol = begin_icol + ifield
            style = field_attr_dict.get('style',default_cell_font)
            val = obj_data.get(fname,None)
            val_func = field_attr_dict.get('val_func',None)
            if val_func:
                try:
                    val = val_func(val,obj_data,FIELDNAME_FIELDATTR)
                except TypeError:
                    val = val_func(val,obj_data,FIELDNAME_FIELDATTR,{'irow':irow, 'icol':icol,'write_a_row_more_data':write_a_row_more_data})
            
            if all_cell_func:
                val = all_cell_func(val)
            field_attr_dict['val'] = val
            is_temp_field = field_attr_dict.get('is_temp_field',False)
            if is_temp_field:
                continue
            ws.write(irow, icol,val, style)
            if height != None:
                ws.row(irow).height_mismatch = True
                ws.row(irow).height = height
            ifield +=1
    default_merge_title_font = table_setting['default_merge_title_font']
    default_title_font = table_setting['default_title_font']
    def write_a_title(FIELDNAME_FIELDATTR, ws, irow, begin_icol, default_width, is_merge_title,merge_title_font=default_merge_title_font,title_font=default_title_font ):
        ifield = 0 
        if is_merge_title:
            merge_title_irow = irow
            title_irow = irow + 1
            if title_height != None:
                ws.row(merge_title_irow).height_mismatch = True
                ws.row(merge_title_irow).height = title_height
        else:
            title_irow = irow
            
            
        if title_height != None:
            ws.row(title_irow).height_mismatch = True
            ws.row(title_irow).height = title_height
        merge_title_old = None
        for fname, field_attr_dict  in FIELDNAME_FIELDATTR.items():
            is_temp_field = field_attr_dict.get('is_temp_field',False)
            if is_temp_field:
                continue
            icol = begin_icol + ifield
            title = field_attr_dict.get('title', fname)
            
            if is_merge_title:
                merge_title = field_attr_dict.get('merge_title', None)
                if merge_title==None or merge_title != merge_title_old:#merge_title and merge_title == merge_title_old
                    ws.write(merge_title_irow,icol,merge_title, title_font)
                    merge_title_icol_old = icol
                else:
                    ws.write_merge(merge_title_irow,merge_title_irow,merge_title_icol_old, icol, merge_title, merge_title_font )
                merge_title_old = merge_title
            if is_merge_title and  merge_title == None:
                ws.write_merge(merge_title_irow, title_irow, icol, icol, title , merge_title_font)
            else:   
                ws.write(title_irow, icol, title , title_font)
            if not skip_width:
                width = field_attr_dict.get('width',None)
                if width:
                    width = get_width(width)
                else:
                    auto_width = field_attr_dict.get('auto_width',False)
                    if auto_width:
                        width = get_width(len(title) )
                    elif default_width:
                        width = get_width(default_width)
                if width:
                    ws.col(icol).width = width
            ifield +=1
        if is_merge_title:
            return 2
        else:
            return 1
        
        
    FIELDNAME_FIELDATTR = table_setting['FIELDNAME_FIELDATTR']
    FIELDNAME_FIELDATTR = OrderedDict (FIELDNAME_FIELDATTR)
    default_width = table_setting.get('default_width',10)
    is_merge_title = table_setting.get('is_merge_title',False)  
    gen_row_data = table_setting.get('gen_row_data')
    title_nrow = write_a_title(FIELDNAME_FIELDATTR, ws, begin_title_irow, begin_icol, default_width, is_merge_title)        
    irow = begin_title_irow +title_nrow   
    write_a_row_more_data = {'begin_iabrow':irow}
    for  i in datas:
        if gen_row_data:
            obj_data = gen_row_data(i)
        else:
            obj_data = i
        write_a_row(obj_data, FIELDNAME_FIELDATTR, ws, irow, begin_icol,write_a_row_more_data=write_a_row_more_data)
        irow +=1
    nrow = title_nrow + len(datas)
    return nrow
# new

def get_variable_values(request_args):
    variable_values = {}
    if 'from' in request_args:
        variable_values['from'] = request_args['from']
    if 'to' in request_args:
        variable_values['to'] = request_args['to']
    return variable_values
        
font_map = {1:'Calibri', 2:'Times New Roman'}
def get_font_font_size(request_args):
    font_font_size_dict = {}
    if 'font_size' in request_args:
        font_size = request_args['font_size']
        if font_size > 9 and font_size < 13:
            font_font_size_dict['font_size'] =  font_size
    if 'font' in request_args:
        font = request_args['font']
        if font in font_map:
            font = font_map.get(font)
            font_font_size_dict['font'] =  font
    return font_font_size_dict

def common_one_table_report_xl(request_args, basic_setting, gen_table_setting_list, gen_fixups):
    font_font_size_dict = get_font_font_size(request_args)
    font = font_font_size_dict.get('font') or  basic_setting['Font_default']
    font_size = font_font_size_dict.get('font_size') or basic_setting['Font_size_default']
   
    table_setting_list =[]
    if not isinstance(gen_table_setting_list, list):
        gen_table_setting_list = [gen_table_setting_list]
    for count, gen_table_setting in enumerate(gen_table_setting_list):
        table_setting = gen_table_setting (font, font_size, request_args)
        if table_setting.get('get_hasura_data',True):
            get_variable_values_func = table_setting.get('get_variable_values', get_variable_values)
            variable_values = get_variable_values_func(request_args)
            data_hasura =  get_hasura_data_with_query_and_variable( variable_values=variable_values, query= table_setting['query'])
        print ('***data_hasura','count', count,  data_hasura)
        out_datas_func=table_setting.get('out_datas_func')
        if out_datas_func:
            datas =out_datas_func(data_hasura)
        else:
            datas = data_hasura
        table_setting['datas'] = datas
#         print ('datas', datas)
        table_setting_list.append(table_setting)
    if len(table_setting_list)==1:
        table_setting_list = table_setting_list[0]
    setting_fixups = gen_fixups(font, font_size,variable_values, table_setting_list, request_args)
    wb = write_fixups(setting_fixups )
    return wb

#usually func
def convert_gmt_str_dt_to_vn_str_dt(from_):
    from_ = datetime.strptime(from_,'%Y-%m-%d')
#     from_ = from_.strftime('%d/%m/%Y')
    return  from_

def display_from_to(variable_values):
    from_ = variable_values['from']
    to_ = variable_values['to']
    if from_:
        variable_values['from'] = from_
        try:
            from_ = datetime.strptime(from_,'%Y-%m-%d')
            from_ = from_.strftime('%d/%m/%Y')
        except ValueError:
            from_ = ''
    if to_:
        variable_values['to'] = to_
        try:
            to_ = datetime.strptime(to_,'%Y-%m-%d')
            to_ = to_.strftime('%d/%m/%Y')
        except ValueError:
            to_ = ''
    return u'Từ ngày %s đến ngày %s'%(from_, to_)

def easyxf_new(str_style,**kargs):
    style = xlwt.easyxf(str_style)
    for k,v in kargs.items():
        setattr(style, k, v)
    return style
    
    
#######################################NN#######################
Basic_setting = {
    'Font_default':'Calibri',
    'Font_size_default':11
    }
def ne_nep_gen_table_setting (font, font_size, request_args):
    ne_nep_query = '''query($from:date!,$to:date!){
  result:edu_classes_aggregate(
    order_by: {
      class_name: asc_nulls_last
    }    
  ){
    aggregate { count }
    nodes {
      id
      key:id
      class_name
      statistics:thong_ke_vi_pham_tap_thes_aggregate(
        where:{      
          day_work:{
            _gte:$from,
            _lte:$to
          }      
        }
      ){
        aggregate{
          sum{          
            slg_loi_tap_the
            diem_tru_tap_the
            slg_loi_ca_nhan
            diem_tru_ca_nhan
            slg_loi_diem_danh
            diem_tru_diem_danh
            tong_diem_tru
          }
        }
      }
    }
  }
}''' 
    
    def ne_nep_out_datas_func (rs):
        datas =rs['data']['result']['nodes']
        new_datas = []
        for d in datas:
            d = ne_nep_gen_row_data(d)
            if d['tong_diem_tru'] ==None:
                d['tong_diem_tru'] =0
                
            new_datas.append(d)
        datas = new_datas
        print ('**datas',datas)
        add_stt_and_xh(datas, sorted_key='tong_diem_tru')
        return datas
    def ne_nep_gen_row_data(data_item):
        class_name  = data_item['class_name']
        obj_data = data_item['statistics']['aggregate']['sum']
        obj_data['class_name'] = class_name
        return obj_data
    

    @font_decorator_parent_new(font = font, height = font_size, vert = 'center',  horiz = 'center' )
    def generate_easyxf(*args, **kargs):
        return generate_easyxf_import(*args, **kargs)
    ne_nep_FIELDNAME_FIELDATTR = [
#             ('stt',{'val_func': lambda v,d,s:s['stt'].get('val',0)+1,'width':4,'title':u'STT'}),
            ('stt',{'title':u'STT', 'width':4}),
            ('xep_hang',{'title':u'XH', 'width':4}),
            ('class_name',{'title':u'Lớp', 'width':5}),
            ('slg_loi_tap_the',{'merge_title':u'Vi phạm tập thể','title':u'Số lỗi vi phạm', 'val_func':convert_none_to_0}),
            ('diem_tru_tap_the',{'merge_title':u'Vi phạm tập thể','title':u'Số điểm trừ' , 'auto_width':True}),
            ('slg_loi_ca_nhan',{ 'merge_title':u'Vi phạm nề nếp cá nhân', 'title':u'Số lỗi vi phạm'}),
            ('diem_tru_ca_nhan',{'merge_title':u'Vi phạm nề nếp cá nhân', 'title':u'Số điểm trừ', 'auto_width':True}),
            ('slg_loi_diem_danh',{'merge_title':u'Vi phạm điểm danh', 'title':u'Số lỗi vi phạm' }),
            ('diem_tru_diem_danh',{'merge_title':u'Vi phạm điểm danh', 'title':u'Số điểm trừ', 'auto_width':True}),
            ('tong_diem_tru',{'title':u'Tổng điểm trừ',  'width':7.6}),
            ]
        
    ne_nep_table_setting = {
            'query':ne_nep_query,
            'out_datas_func':ne_nep_out_datas_func,
            'gen_row_data':None,
            'default_width':11,
            'default_merge_title_font':xlwt.easyxf(generate_easyxf(align_wrap=True, bold=True,vert = 'center',horiz='center',borders='left thin, right thin, top thin, bottom thin')),
            'default_title_font':xlwt.easyxf(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center',horiz = 'center')),
            'default_cell_font':xlwt.easyxf(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center')),
            'is_merge_title':True,
            'FIELDNAME_FIELDATTR':ne_nep_FIELDNAME_FIELDATTR,
            'row_height':20*16,
            'title_height':20*26,
            'all_cell_func':convert_none_to_0,
            }
    return ne_nep_table_setting

def ne_nep_gen_fixups(font, font_size,variable_values, ne_nep_table_setting, request_args):
    @font_decorator_parent_new(font = font, height = font_size, vert = 'center',  horiz = 'center' )
    def generate_easyxf(*args, **kargs):
        return generate_easyxf_import(*args, **kargs)
    Begin_irow = 2
    Begin_icol = 0
    
    fixups_ne_nep =[  
         ('header',{'range':[Begin_irow, Begin_irow, Begin_icol, Begin_icol+7],'val':u'BÁO CÁO NỀ NẾP VI PHẠM TỔNG HỢP', 
                    'style':xlwt.easyxf(generate_easyxf(bold=True)), 'row_height': 20*21}),
         ('header2',{'range':['auto', 0, Begin_icol, Begin_icol+7],'val_func':display_from_to, 'val_func_kargs':{'variable_values':variable_values},
                      'style':xlwt.easyxf(generate_easyxf(bold=True)),  'row_height': 20*21}),
         ('table',{'range':['auto', Begin_icol],'val':None,
               'func':write_table_rerange, 'offset':2,
               'func_kargs':{'table_setting':ne_nep_table_setting,
                                }  }),
#          ('header3',{'range':['auto',Begin_icol],'val':u'Kết thúc'}),
                     ] 
    setting_fixups = {
        'row_height':20*16, 
        'fixups':fixups_ne_nep,
        'default_fixups_style':xlwt.easyxf(generate_easyxf(vert = 'center',horiz = 'center'))
        }
    return setting_fixups


def ne_nep_report_xl(request_args):
    wb = common_one_table_report_xl(request_args, Basic_setting, [ne_nep_gen_table_setting], ne_nep_gen_fixups)
    return wb

######################################END NN###########################

#############################DIEM DANH ###########################
Basic_setting = {
    'Font_default':'Calibri',
    'Font_size_default':11

    }
def diem_danh_gen_table_setting (font, font_size, request_args):
    diem_danh_query = '''query($from:date!,$to:date!){
  result:edu_student_attendances(
    where:{   
      _and:[{
        attend_date:{
          _gte:$from,
          _lte:$to
        }
      }, {
        active: {
          _eq:true
        }
      }]               
    },
    order_by: {
      class_enrollment: {
        class: {
          class_name: asc_nulls_last
        },
        student:{
          last_name:asc_nulls_last,
          first_name:asc_nulls_last
        }
      }      
    }
  ){   
    id
    key:id
    class_enrollment {
      id
      class{
        id
        class_name
      }
      student {
        id
        student_code
        first_name
        last_name
      }
    }
    attend_date
    attendance_type {
      id
      description
    }
  }
}''' 
    
    def diem_danh_out_datas_func (rs):
        datas =rs['data']['result']
#         add_stt_and_xh(datas, sorted_key = 'so_diem_tru') 
#         print ('**datas_diem_danh', datas_diem_danh)

        return datas


    @font_decorator_parent_new(font = font, height = font_size )
    def generate_easyxf(*args, **kargs):
        return generate_easyxf_import(*args, **kargs)
    
    diem_danh_FIELDNAME_FIELDATTR = [
            ('stt',{'val_func': lambda v,d,s:s['stt'].get('val',0)+1,'width':6, 'title':'STT'}),
#             ('stt',{'title':u'STT', 'width':4}),
#             ('xep_hang',{'title':u'XH', 'width':4}),
            ('student_code',{'val_func': lambda v,data_obj, sum_data: data_obj['class_enrollment']['student']['student_code'],  'width':13, 'title':u'Mã học sinh'}),
            ('student_name',{'title':u'Họ và tên', 'width':25,'val_func': lambda v,data_obj,sum_data: data_obj['class_enrollment']['student']['first_name'] +' ' +  data_obj['class_enrollment']['student']['last_name'] }),      
            ('class_name',{'val_func': lambda v,data_obj, sum_data: data_obj['class_enrollment']['class']['class_name'], 'width':6, 'title':'Lớp học'}),
            ('attend_date',{'title':u'Ngày vi phạm', 'val_func': lambda v,*args: convert_gmt_str_dt_to_vn_str_dt(v[0:10]),
                            'style':easyxf_new(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center'),num_format_str = 'dd/mm/yyyy'), 
                            }),
            ('attendance_type',{'title':u'Lỗi vi phạm', 'val_func': lambda v,d,s: v['description'], 'width':25}),
            ('so_diem_tru',{'title':u'Số điểm trừ', 'width':11}),
            ]
        
    diem_danh_table_setting = {
            'query':diem_danh_query,
            'out_datas_func':diem_danh_out_datas_func,
            'gen_row_data':None,
            'default_width':15,
            'default_merge_title_font':xlwt.easyxf(generate_easyxf(bold=True,vert = 'center',horiz='center',borders='left thin, right thin, top thin, bottom thin')),
            'default_title_font':xlwt.easyxf(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center',horiz = 'center')),
            'default_cell_font':xlwt.easyxf(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center')),
            'is_merge_title':True,
            'FIELDNAME_FIELDATTR':diem_danh_FIELDNAME_FIELDATTR,
            'row_height':20*16,
            'title_height':20*26,
            'all_cell_func':convert_none_to_0,
            }
    return diem_danh_table_setting
def convert_none_to_0(v,*args):
        if v == None:
            v = 0
        return v
def diem_danh_gen_fixups(font, font_size,variable_values, diem_danh_table_setting,request_args):
    print ('***variable_values', variable_values)
    @font_decorator_parent_new(font = font, height = font_size, vert = 'center',horiz = 'center')
    def generate_easyxf(*args, **kargs):
        return generate_easyxf_import(*args, **kargs)
    Begin_irow_diem_danh = 2
    Begin_icol_diem_danh = 0
    fixups_diem_danh =[  
     ('header',{'range':[Begin_irow_diem_danh, Begin_irow_diem_danh, Begin_icol_diem_danh, Begin_icol_diem_danh+7], 'val':u'THỐNG KÊ ĐIỂM DANH', 
                'style':xlwt.easyxf(generate_easyxf(bold=True)),
                'row_height': 20*21
                }),
     ('header2',{'range':['auto', 0, Begin_icol_diem_danh,Begin_icol_diem_danh+7],'val_func':display_from_to, 'val_func_kargs':{'variable_values':variable_values},
                    'style':xlwt.easyxf(generate_easyxf(bold=True)),
                  }),
     ('table',{'range':['auto', Begin_icol_diem_danh],'val':None,
               'func':write_table_rerange, 'offset':2,
               'func_kargs':{'table_setting':diem_danh_table_setting,
                                }  }),
                     ] 
    
    setting_fixups = {
        'fixups':fixups_diem_danh,
        'default_fixups_style':xlwt.easyxf(generate_easyxf(vert = 'center',horiz = 'center'))
        }
    return setting_fixups


def diem_danh_report_xl(request_args):
    wb = common_one_table_report_xl(request_args, Basic_setting, diem_danh_gen_table_setting, diem_danh_gen_fixups)
    return wb

#############################!DIEM DANH ###########################


#############################NN CHITIET ###########################

Basic_setting = {
    'Font_default':'Calibri',
    'Font_size_default':11
    }
Default_break_sheet = True

def nn_chi_tiet_ca_nhan_gen_table_setting (font, font_size,request_args):
    ne_nep_query = '''query($from:date!,$to:date!){
  result1:thong_ke_truong_v_thongke_bao_cao_vi_pham_ne_nep_chi_tiet_aggregate(
    order_by: {
      class_name: asc_nulls_last
    },
    where: {
      _and: {
        attend_date:{
          _gte:$from,
          _lte:$to
        }        
      }      
    } 
  ){
    aggregate { count }
    nodes {
      violated_at:attend_date
      class_name
      first_name
      last_name
      ten_vi_pham_ne_nep:loi_vi_pham
      punish_point
    }
  }  
  result2:thong_ke_truong_v_thongke_vipham_ne_nep_tap_the_chi_tiet_aggregate(
    order_by: {
      class_name: asc_nulls_last
    },
    where: {
      _and: {
        violation_date:{
          _gte:$from,
          _lte:$to
        }        
      }      
    } 
  ){
    aggregate { count }
    nodes {
      violation_date
      class_name
         violation_name     
      punish_point
    }
  }
}''' 
    
#     def get_variable_values_nn_chi_tiet(request_args):
#         variable_values = {}
#         if 'from' in request_args:
#             variable_values['from'] = request_args['from']
#             variable_values['from1'] = request_args['from']
#         if 'to' in request_args:
#             variable_values['to'] = request_args['to']
#             variable_values['to1'] = request_args['to']
#         return variable_values

    def nn_chi_tiet_out_datas_func (rs):
        datas =rs['data']['result1']['nodes']
        add_stt_and_xh(datas, sorted_key = 'punish_point')    
        
#         datas.sort(key =lambda x: -x['punish_point'])
#         for count, i in enumerate(datas):
#             i['xep_hang'] = count +1
#         sorted_datas = sorted(datas, lambda x: x['punish_point'])
        return datas
    @font_decorator_parent_new(font = font, height = font_size, vert = 'center',  horiz = None )
    def generate_easyxf(*args, **kargs):
        return generate_easyxf_import(*args, **kargs)
    
    def xep_hang_func(v,d,s,*arg):
        xh_cu = s['xep_hang'].get('val',1)
        punish_point_cu = s['punish_point']
        punish_point = d['punish_point']
        if punish_point == punish_point_cu:
            xh = xh_cu
        else:
            xh = xh_cu +1
        return xh
    def formula_write(v,d,s,more_data):
        irow = more_data['irow'] + 1
        icol = more_data['icol']
        iabcol = Convert_number_to_APB[icol]
        write_a_row_more_data = more_data['write_a_row_more_data']
        begin_iabrow = write_a_row_more_data['begin_iabrow'] + 1
        formula = 'RANK(G%s,$G$%s:$G$108)'%(irow,begin_iabrow)
#         formula = 'rank(%(iabcol)s%(irow)s,$(iabcol)s$8:$(iabcol)s$108)'%{'iabcol':iabcol, 'irow':irow}
#         formula = 'rank(%s(iabcol)%s(irow),$s(iabcol)$8:$s(iabcol)$108)'%{'iabcol':iabcol, 'irow':irow}

        return xlwt.Formula(formula)
    ne_nep_FIELDNAME_FIELDATTR = [
#             ('stt',{'val_func': lambda v,d,s:s['stt'].get('val',0)+1,'width':4,'title':u'STT'}),

            ('stt',{'title':u'STT'}),
            ('xep_hang',{'title':u'XH' , 'width':4}),
            ('violated_at',{'val_func': lambda v,*args: convert_gmt_str_dt_to_vn_str_dt(v[0:10]), 
                            'style':easyxf_new(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center'),num_format_str = 'dd/mm/yyyy'), 
                            'title':'Ngày vi phạm'}),
            ('class_name',{'title':u'Lớp', 'width':5}),
            ('first_name',{'title':u'Học sinh vi phạm','val_func': lambda v,d,s:v +' '  + d['last_name'], 'width':35}),
            ('ten_vi_pham_ne_nep',{'title':u'Lỗi vi phạm','width':25}),
            ('punish_point',{'title':u'Số điểm trừ', 'width':11}),
            ]
        
    ne_nep_table_setting = {
#             'get_variable_values':get_variable_values_nn_chi_tiet,
            'query':ne_nep_query,
            'out_datas_func':nn_chi_tiet_out_datas_func,
            'gen_row_data':None,
            'default_width':11,
            'default_merge_title_font':xlwt.easyxf(generate_easyxf(align_wrap=True, bold=True,vert = 'center',horiz='center',borders='left thin, right thin, top thin, bottom thin')),
            'default_title_font':xlwt.easyxf(generate_easyxf(bold = True, borders='left thin, right thin, top thin, bottom thin',vert = 'center',horiz = 'center', pattern = 'pattern solid, fore_colour gray25')),
            'default_cell_font':xlwt.easyxf(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center')),
            'is_merge_title':False,
            'FIELDNAME_FIELDATTR':ne_nep_FIELDNAME_FIELDATTR,
            'row_height':20*16,
            'title_height':20*26,
            'all_cell_func':convert_none_to_0,
            }
    return ne_nep_table_setting

def add_xh(datas, sorted_key ='punish_point'):
    xh_old = 0 
    dong_xep_hang = 0
    for count, d in enumerate(datas):
        punish_point = d[sorted_key]
        if count ==0 or punish_point != punish_point_old:
            xh = xh_old  + dong_xep_hang + 1
            dong_xep_hang = 0 
        else:
            xh = xh_old
            dong_xep_hang+=1
        xh_old = xh
        punish_point_old = punish_point
        d['xep_hang'] = xh
    return datas
        
def add_stt_and_xh(datas, sorted_key = 'punish_point',item_key = None):
    for count, i in enumerate(datas):
        i['stt'] = count +1 
    item_key = item_key or sorted_key
    datas = sorted(datas, key=itemgetter(item_key))
#     datas = sorted(datas, key=lambda x: x[item_key] or 0)

    add_xh(datas, sorted_key =sorted_key)     
    datas = sorted(datas, key=itemgetter('stt'))    
    

def nn_chi_tiet_tap_the_gen_table_setting (font, font_size, request_args):
    break_sheet = request_args.get('break_sheet', Default_break_sheet)
    ne_nep_query = '''query($from1:timestamptz!,$to1:timestamptz!,$from:date!,$to:date!){
  result1:thong_ke_truong_v_thongke_bao_cao_vi_pham_ne_nep_chi_tiet_aggregate(
    order_by: {
      class_name: asc_nulls_last
    },
    where: {
      _and: {
        attend_date:{
          _gte:$from1,
          _lte:$to1
        }        
      }      
    } 
  ){
    aggregate { count }
    nodes {
      violated_at:attend_date
      class_name
      first_name
      last_name
      ten_vi_pham_ne_nep:loi_vi_pham
      punish_point
    }
  }  
  result2:thong_ke_truong_v_thongke_vipham_ne_nep_tap_the_chi_tiet_aggregate(
    order_by: {
      class_name: asc_nulls_last
    },
    where: {
      _and: {
        violation_date:{
          _gte:$from,
          _lte:$to
        }        
      }      
    } 
  ){
    aggregate { count }
    nodes {
      violation_date
      class_name
         violation_name     
      punish_point
    }
  }
}''' 
    
    def ne_nep_tap_the_out_datas_func (rs):
        datas =rs['data']['result2']['nodes']
        add_stt_and_xh(datas,'punish_point')
        return datas

    @font_decorator_parent_new(font = font, height = font_size, vert = 'center',  horiz = None )
    def generate_easyxf(*args, **kargs):
        return generate_easyxf_import(*args, **kargs)
    ne_nep_tap_the_FIELDNAME_FIELDATTR = [
#             ('stt',{'val_func': lambda v,d,s:s['stt'].get('val',0)+1,'width':4,'title':u'STT'}),
            ('stt',{'title':u'STT', 'width':4}),
            ('xep_hang',{'title':u'XH', 'width':4}),
            ('violation_date',{'val_func': lambda v,*args: convert_gmt_str_dt_to_vn_str_dt(v[0:10]),
                            'style':easyxf_new(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center'),num_format_str = 'dd/mm/yyyy'), 
                            'title':u'Ngày vi phạm'}),
            ('class_name',{'title':u'Lớp', 'width':5}),
            ('violation_name',{'title':u'Lỗi vi phạm', 'width':25}),
            ('punish_point',{'title':u'Số điểm trừ', 'width':25}),
            ('ghi_chu',{'title':u'Ghi chú'}),
            ]
        
    table_setting = {
            'skip_width': not break_sheet,
            'get_hasura_data':None,
            'query':ne_nep_query,
            'out_datas_func':ne_nep_tap_the_out_datas_func,
            'gen_row_data':None,
            'default_width':11,
            'default_merge_title_font':xlwt.easyxf(generate_easyxf(align_wrap=True, bold=True,vert = 'center',horiz='center',borders='left thin, right thin, top thin, bottom thin')),
            'default_title_font':xlwt.easyxf(generate_easyxf(bold = True, borders='left thin, right thin, top thin, bottom thin',vert = 'center',horiz = 'center', pattern = 'pattern solid, fore_colour gray25')),
            'default_cell_font':xlwt.easyxf(generate_easyxf(borders='left thin, right thin, top thin, bottom thin', vert = 'center')),
            'is_merge_title':False,
            'FIELDNAME_FIELDATTR':ne_nep_tap_the_FIELDNAME_FIELDATTR,
            'row_height':20*16,
            'title_height':20*26,
            'all_cell_func':convert_none_to_0,
            }
    return table_setting
def nn_chi_tiet_gen_fixups(font, font_size,variable_values, table_setting_list, request_args):
    break_sheet = request_args.get('break_sheet', Default_break_sheet)
    @font_decorator_parent_new(font = font, height = font_size, vert = 'center', horiz = 'center' )
    def generate_easyxf(*args, **kargs):
        return generate_easyxf_import(*args, **kargs)
    Begin_irow = 2
    Begin_icol = 0
    
    fixups_ne_nep =[  
        ('header',{'range':[Begin_irow, Begin_irow, Begin_icol, Begin_icol+7], 'val': u'BÁO CÁO NỀ NẾP VI PHẠM TỔNG HỢP', 
                    'style':xlwt.easyxf(generate_easyxf(bold=True)),
                    'row_height': 20*21
                    }),
        ('header2',{'range':['auto', 0, Begin_icol, Begin_icol+7],'val_func':display_from_to, 'val_func_kargs':{'variable_values':variable_values},
                      'style':xlwt.easyxf(generate_easyxf(bold=True)),
#                       'row_height': 20*21
                      }),
        ('header3',{'range':['auto', 0, Begin_icol, Begin_icol+7], 'val':u'I. Vi phạm cá nhân', 'style':xlwt.easyxf(generate_easyxf(bold=True, horiz = 'left'))
                   }),          
                    
        ('table',{'range':['auto', Begin_icol],'val':None,
               'func':write_table_rerange, 'offset':2,
               'func_kargs':{'table_setting':table_setting_list[0],
                                }  }),
        
        ('header_break',{'skip_row':not break_sheet, 'break_sheet':break_sheet, 'sheet_name': u'Chi tiết tập thể' , 'range':[Begin_irow, Begin_irow, Begin_icol, Begin_icol+7],'val':u'BÁO CÁO NỀ NẾP VI PHẠM TỔNG HỢP', 
                    'style':xlwt.easyxf(generate_easyxf(bold=True)), 'row_height': 20*21}),
        ('header_break2',{'skip_row':not break_sheet, 'range':['auto', 0, Begin_icol, Begin_icol+7],'val_func':display_from_to, 'val_func_kargs':{'variable_values':variable_values},
                      'style':xlwt.easyxf(generate_easyxf(bold=True)),  'row_height': 20*21}),
                            
        ('header4',{'range':['auto',0, Begin_icol, Begin_icol+7], 'val':u'II. Vi phạm tập thể', 'style':xlwt.easyxf(generate_easyxf(bold=True, horiz = 'left'))
                   }),  
                    
        ('table2',{'range':['auto', Begin_icol],'val':None,
                     'func':write_table_rerange, 'offset':2,
                     'func_kargs':{'table_setting':table_setting_list[1],
                          }  }),
                     ] 
    setting_fixups = {
        'sheet_name': u'Chi tiết cá nhân và tập thể' if not break_sheet else u'Chi tiết cá nhân' , 
        'row_height':20*16, 
        'fixups':fixups_ne_nep,
        'default_fixups_style':xlwt.easyxf(generate_easyxf(vert = 'center',horiz = 'center'))
        }
    return setting_fixups


def nn_chi_tiet_report_xl(request_args):
    wb = common_one_table_report_xl(request_args, Basic_setting, [nn_chi_tiet_ca_nhan_gen_table_setting, nn_chi_tiet_tap_the_gen_table_setting], nn_chi_tiet_gen_fixups)
    return wb
#############################!NN CHITIET ###########################


Convert_dict = {'false':False, 'true':True, '^\d+$': lambda v:int(v), '^\d+\.(\d*)$': lambda v:float(v)}
def convert_type(request_args):
    new_kargs = {}
    for k_rq, v in request_args.items():
        if isinstance(v, str):
            for pt, repl in Convert_dict.items():
                is_match = re.search(pt, v,re.I)
                print (k_rq, pt,v,is_match)
                if is_match:
                    if callable(repl):
                        v = repl(v)
                    else:
                        v = repl
                    break
        new_kargs[k_rq]= v
    return new_kargs
dlxl_map_func = {'ne_nep':{'func':ne_nep_report_xl,'file_name':'ne_nep_tong_hop'},
                'diem_danh':{'func':diem_danh_report_xl,'file_name':'vi_pham_diem_danh'},
                'nn_chi_tiet':{'func':nn_chi_tiet_report_xl,'file_name':'nn_chi_tiet'},
                 }

def get_funcxl_and_run_funcxl_from_key(func_key, request_args):
#     print ('**request_args', request_args)
    request_args = convert_type(request_args)
#     print ('**request_args convert type', request_args)
    adict = dlxl_map_func[func_key]
    func = adict['func']
    filename = adict['file_name'] +'.xls'
    wb = func(request_args)
    print ('done gen file')
    return wb, filename

@app.route('/')
@app.route('/index')
def index():
    return "Hello, World tu day"
    
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




if __name__ == '__main__':
    dlxl_map_func = {'ne_nep':{'func':ne_nep_report_xl,'file_name':'ne_nep_tong_hop'},
                'diem_danh':{'func':diem_danh_report_xl,'file_name':'vi_pham_diem_danh'},
                'nn_chi_tiet':{'func':nn_chi_tiet_report_xl,'file_name':'nn_chi_tiet'},
                 }
    
    variable_values_dd ={ "from": "2019-09-14", "to": "2019-09-16", 'break_sheet': False  }  
#     wb = nn_chi_tiet_report_xl(variable_values_dd)
    key = 'ne_nep'
    adict = dlxl_map_func[key]
    wb =adict ['func'](variable_values_dd)
    wb.save(r'C:\Users\tu\Desktop\New folder\%s_%s.xls'%(adict['file_name'],datetime.now().strftime('%d_%m_%H_%M_%S')))       
    print('done')




