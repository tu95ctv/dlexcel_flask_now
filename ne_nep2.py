# -*- coding: utf-8 -*-
import xlwt
import requests
from collections import  OrderedDict
from time import sleep
from datetime import datetime
# Font = 'Calibri'

def font_decorator_parent(font,height):
    def font_decorator(func):
        def awrapper(*args, **kargs):
            if kargs.get('font')==None:
                kargs['font'] = font
            if kargs.get('height')==None:
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

# header_bold_style = xlwt.easyxf(generate_easyxf(height=12,bold=True,vert = 'center',horiz='center',borders='left thin, right thin, top thin, bottom thin',pattern = 'pattern solid, fore_colour gray25'))
# header_bold_style_align_wrap = xlwt.easyxf(generate_easyxf(align_wrap = True, height=12,bold=True,vert = 'center',horiz='center',borders='left thin, right thin, top thin, bottom thin',pattern = 'pattern solid, fore_colour gray25'))
# header_bold_style_no_gray = xlwt.easyxf(generate_easyxf(height=12,bold=True,vert = 'center',horiz='center',borders='left thin, right thin, top thin, bottom thin'))
# header_bold_style_no_gray_no_bottom = xlwt.easyxf(generate_easyxf(height=12,bold=True,vert = 'center',horiz='center',borders='left thin, right thin, top thin'))
# center_border_style = xlwt.easyxf(generate_easyxf(height=12,borders='left thin, right thin, top thin, bottom thin',vert = 'center',horiz = 'center'))
# nocenter_border_style = xlwt.easyxf(generate_easyxf(height=12,borders='left thin, right thin, top thin, bottom thin',vert = 'center'))

generate_easyxf_import = generate_easyxf

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
            ws.write(irow, icol,  obj_data[fname], style)
            
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
            val = v_fixups['val']
            style = v_fixups.get('style',defaut_fixups_style)
            if len(xrange) == 2:
                ws.write(xrange[0], xrange[1], val, style)
            else:
                ws.write_merge(xrange[0], xrange[1],xrange[2], xrange[3], val, style)
    return wb
    
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
# 


def ne_nep_out_datas_func (rs):
    datas_ne_nep =rs['data']['result']['nodes']
    return datas_ne_nep

def write_hasura_table(ws, begin_title_irow, begin_icol, table_setting=None, out_datas_func=None, gen_row_data=None,  query=None, variable_values=None):
        ne_nep_query_data = {'query': query}
        if variable_values:
            ne_nep_query_data['variables'] = variable_values
        rs = get_hasura_data(ne_nep_query_data)
        if out_datas_func:
            datas =out_datas_func(rs)
        else:
            datas = rs
        nrow = write_table(table_setting, datas, begin_title_irow, begin_icol, ws, gen_row_data= gen_row_data)
        return nrow



def ne_nep_gen_row_data(data_item):
    class_name  = data_item['class_name']
    obj_data = data_item['statistics']['aggregate']['sum']
    obj_data['class_name'] = class_name
    return obj_data
    
    
    

#     
# def report_hasura(fixups, default_fixups_style =None):
#     wb = write_fixups(fixups, defaut_fixups_style = default_fixups_style )
#     return wb
#     
    
def ne_nep_report_xl(variable_values):
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

    ne_nep_height_font = 11 
    @font_decorator_parent('Calibri',ne_nep_height_font)
    def generate_easyxf(*args, **kargs):
        return generate_easyxf_import(*args, **kargs)
    FIELDNAME_FIELDATTR_ne_nep = [
            ('class_name',{'title':u'Lớp', 'auto_width':True}),
            ('slg_loi_tap_the',{'merge_title':u'Vi phạm tập thể','title':u'Số lỗi vi phạm'}),
            ('diem_tru_tap_the',{'merge_title':u'Vi phạm tập thể','title':u'Số điểm trừ' }),
            ('slg_loi_ca_nhan',{ 'merge_title':u'Vi phạm nề nếp cá nhân', 'title':u'Số lỗi vi phạm'}),
            ('diem_tru_ca_nhan',{'merge_title':u'Vi phạm nề nếp cá nhân', 'title':u'Số điểm trừ'}),
            ('slg_loi_diem_danh',{'merge_title':u'Vi phạm điểm danh', 'title':u'Số lỗi vi phạm' }),
            ('diem_tru_diem_danh',{'merge_title':u'Vi phạm điểm danh', 'title':u'Số điểm trừ'}),
            ('tong_diem_tru',{'width':20,  'title':u'Tổng điểm trừ'}),
            ]
    ne_nep_table_setting = {
            'default_width':13,
            'default_merge_title_font':xlwt.easyxf(generate_easyxf(bold=True,vert = 'center',horiz='center',borders='left thin, right thin, top thin, bottom thin')),
            'default_title_font':xlwt.easyxf(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center',horiz = 'center')),
            'default_cell_font':xlwt.easyxf(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center',horiz = 'center')),
            'is_merge_title':True,
            'FIELDNAME_FIELDATTR':FIELDNAME_FIELDATTR_ne_nep
            }
    Begin_irow_ne_nep = 2
    Begin_icol_ne_nep = 0
    
    
    
    fixups_ne_nep =[  
 ('header',{'range':[Begin_irow_ne_nep, Begin_irow_ne_nep, Begin_icol_ne_nep, Begin_icol_ne_nep+7],'val':u'Bảng nề nếp', 'style':xlwt.easyxf(generate_easyxf(bold=True,height=20, vert = 'center',horiz = 'center'))}),
 ('header2',{'range':['auto', 0, Begin_icol_ne_nep,Begin_icol_ne_nep+7],'val':u'Từ ngày %s đến ngày %s'%(from_,to_), 'style':xlwt.easyxf(generate_easyxf( vert = 'center',horiz = 'center'))}),
 ('table',{'range':['auto', Begin_icol_ne_nep],'val':None,
           'func':write_hasura_table, 'offset':2,
           'func_kargs':{'query': ne_nep_query,'variable_values':variable_values, 'table_setting':ne_nep_table_setting,
                         'out_datas_func':ne_nep_out_datas_func,
                         'gen_row_data':ne_nep_gen_row_data }  }),
 ('header3',{'range':['auto',Begin_icol_ne_nep],'val':u'Kết thúc'}),
                     ] 

    default_fixups_style  =  xlwt.easyxf(generate_easyxf(height=ne_nep_height_font,vert = 'center',horiz = 'center'))
    wb = write_fixups(fixups_ne_nep, defaut_fixups_style = default_fixups_style )
    return wb



# variable_values ={ "from": "1999-01-01", "to": "2019-10-10" }  
# wb = ne_nep_report_xl(variable_values)
# wb.save(r'C:\Users\tu\Desktop\New folder\abc.xls')       

print('done')


