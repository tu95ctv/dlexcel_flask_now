# -*- coding: utf-8 -*-
import requests
from collections import  OrderedDict
from time import sleep
import xlwt
from datetime import datetime
font_map = {1:'Calibri', 2:'Times New Roman'}
def font_decorator_parent(font, height):
    def font_decorator(func):
        def awrapper(*args, **kargs):
            if 'font' not in kargs and font:
                try:
                    font_style = int(font)
                    font_style = font_map.get(font_style)
                except:
                    font_style = font
                kargs['font'] = font_style
            if 'height' not in kargs and height:
                kargs['height'] = height
            rs = func(*args, **kargs)
            return rs
        return awrapper
    return font_decorator

def generate_easyxf (font='Times New Roman', 
                     bold = False,
                     underline=False,
                     height=12, 
                     align_wrap = False,
                     vert = False,
                     horiz = False,
                     borders = False,
                     pattern = False,
                     italic= False
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
        
    sums = ';'.join(sums)   
    
    return sums


def get_hasura_data(query_data):
    url = 'https://qlth.hpz.vn/v1/graphql'
    headers = {'x-hasura-admin-secret': 'hpz', 'content-type': 'application/json', 'User-Agent' : 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.101 Safari/537.36'}
    count_fail = 0
    while 1:
        print ('get html',url)
        try:
            request = requests.post(url, json=query_data, headers=headers)
            return request.json()
        except Exception as e:
            count_fail +=1
            print ('loi khi get html',e)
            sleep(5)
            if count_fail ==5:
                raise ValueError(u'Lỗi get html')

def write_table(table_setting, datas, begin_title_irow,  begin_icol, ws,gen_row_data = None ):
    
     
    def get_width(num_characters,font_height=12):
        return int((1+num_characters) * 256*font_height/12)
    
    default_cell_font = table_setting['default_cell_font']
    def write_a_row(obj_data, FIELDNAME_FIELDATTR,  ws, irow, begin_icol,cell_font=default_cell_font):
        
        ifield = 0 
        for fname, field_attr_dict  in FIELDNAME_FIELDATTR.items():
            icol = begin_icol + ifield
            style = field_attr_dict.get('style',cell_font)
            val = obj_data.get(fname,None)
            val_func = field_attr_dict.get('val_func',None)
            if val_func:
                val = val_func(val,obj_data,FIELDNAME_FIELDATTR)
            field_attr_dict['val'] = val
            is_temp_field = field_attr_dict.get('is_temp_field',False)
            if is_temp_field:
                continue
            ws.write(irow, icol,val, style)
            
            ifield +=1
    default_merge_title_font = table_setting['default_merge_title_font']
    default_title_font = table_setting['default_title_font']
    def write_a_title(FIELDNAME_FIELDATTR, ws, irow, begin_icol, default_width, is_merge_title,merge_title_font=default_merge_title_font,title_font=default_title_font ):
#         merge_title_font = header_bold_style_no_gray
#         title_font = center_border_style
        
        
        ifield = 0 
        if is_merge_title:
            merge_title_irow = irow
            title_irow = irow + 1
        else:
            title_irow = irow
        merge_title_old = None
    #     merge_title_icol_old = begin_icol
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
            width = field_attr_dict.get('width',None)
            
            
            if width:
                width = get_width(width)
            else:
                auto_width = field_attr_dict.get('auto_width',False)
                if auto_width:
                    width = get_width(len(title) + 4)
                else:
                    width = get_width(default_width)
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
    title_nrow = write_a_title(FIELDNAME_FIELDATTR, ws, begin_title_irow, begin_icol, default_width, is_merge_title)        
    irow = begin_title_irow +title_nrow   
    for  i in datas:
        if gen_row_data:
            obj_data = gen_row_data(i)
        else:
            obj_data = i
        write_a_row(obj_data, FIELDNAME_FIELDATTR, ws, irow, begin_icol)
        irow +=1
    nrow = title_nrow + len(datas)
    return nrow



def write_fixups(fixups, defaut_fixups_style = None):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('A Test Sheet', cell_overwrite_ok=True)
    irow = 0
    fixups = OrderedDict (fixups)
    for k_fixups, v_fixups in fixups.items():
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
    return wb
    

def write_hasura_table(ws, begin_title_irow, begin_icol, table_setting=None, out_datas_func=None, gen_row_data=None,  query=None, variable_values=None):
        ne_nep_query_data = {'query': query}
        if variable_values:
            ne_nep_query_data['variables'] = variable_values
        rs = get_hasura_data(ne_nep_query_data)
#         print (rs)
#         raise ValueError('akakaka')
        if out_datas_func:
            datas =out_datas_func(rs)
        else:
            datas = rs
            
        print (datas)
#         raise ValueError('akakaka')
        nrow = write_table(table_setting, datas, begin_title_irow, begin_icol, ws, gen_row_data= gen_row_data)
        return nrow
    



#usually func

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


    
    