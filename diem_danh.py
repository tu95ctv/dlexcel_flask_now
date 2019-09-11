# -*- coding: utf-8 -*-
import xlwt
from datetime import datetime
from dltool import generate_easyxf as generate_easyxf_import
from dltool import font_decorator_parent, write_hasura_table,write_fixups

from dltool import display_from_to


def diem_danh_out_datas_func (rs):
    datas_diem_danh =rs['data']['result']
    return datas_diem_danh

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


def gen_generate_easyxf(font, font_size):
    def generate_easyxf(*args, **kargs):
        print ('kargs**1', kargs)
        if 'font' not in kargs and font:
            kargs['font'] = font
        if 'height' not in kargs and font_size:
            kargs['height'] = font_size
        print ('kargs**2', kargs)
        return generate_easyxf_import(*args, **kargs)
    return generate_easyxf

def diem_danh_report_xl(variable_values, font_size = 11, font = 1):
#     font_map = {'1':'Calibri', '2':'Times New Roman'}
#     font = font_map[font]
    @font_decorator_parent(font, font_size)
    def generate_easyxf(*args, **kargs):
        return generate_easyxf_import(*args, **kargs)
#     if font_size:
#         @font_decorator_parent('Calibri', font_size)
#         def generate_easyxf(*args, **kargs):
#             return generate_easyxf_import(*args, **kargs)
#     else:
#         def generate_easyxf(*args, **kargs):
#             return generate_easyxf_import(*args, **kargs)
        
    
    
    
    
    FIELDNAME_FIELDATTR_diem_danh = [
            ('stt',{'val_func': lambda v,d,s:s['stt'].get('val',0)+1,'width':6}),
#             ('class_enrollment',{'title':u'Lớp', 'auto_width':True, 'val_func': lambda v,d,s: str(v)}),#,'is_temp_field':True
            ('class_name',{'val_func': lambda v,data_obj, sum_data: data_obj['class_enrollment']['class']['class_name']}),
            ('student_code',{'val_func': lambda v,data_obj, sum_data: data_obj['class_enrollment']['student']['student_code']}),
#             ('student',{'val_func': lambda v,data_obj,sum_data: str(sum_data['class_enrollment']['val']['student'])}),
#             ('student',{'val_func': lambda v,data_obj,sum_data: str(data_obj['class_enrollment']['student'])}), 
            ('student_name',{'width':25,'val_func': lambda v,data_obj,sum_data: data_obj['class_enrollment']['student']['first_name'] + data_obj['class_enrollment']['student']['last_name'] }),      
            ('attend_date',{}),
            ('attendance_type',{'val_func': lambda v,d,s: v['description']}),

            ]
    diem_danh_table_setting = {
            'default_width':15,
            'default_merge_title_font':xlwt.easyxf(generate_easyxf(bold=True,vert = 'center',horiz='center',borders='left thin, right thin, top thin, bottom thin')),
            'default_title_font':xlwt.easyxf(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center',horiz = 'center')),
            'default_cell_font':xlwt.easyxf(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center')),
            'is_merge_title':True,
            'FIELDNAME_FIELDATTR':FIELDNAME_FIELDATTR_diem_danh
            }
    Begin_irow_diem_danh = 2
    Begin_icol_diem_danh = 0
    
    
    
    fixups_diem_danh =[  
     ('header',{'range':[Begin_irow_diem_danh, Begin_irow_diem_danh, Begin_icol_diem_danh, Begin_icol_diem_danh+7],'val':u'Bảng Điểm Danh', 'style':xlwt.easyxf(generate_easyxf(bold=True,height=20, vert = 'center',horiz = 'center'))}),
     ('header2',{'range':['auto', 0, Begin_icol_diem_danh,Begin_icol_diem_danh+7],'val_func':display_from_to, 'val_func_kargs':{'variable_values':variable_values}, 'style':xlwt.easyxf(generate_easyxf( vert = 'center',horiz = 'center'))}),
     ('table',{'range':['auto', Begin_icol_diem_danh],'val':None,
               'func':write_hasura_table, 'offset':2,
               'func_kargs':{'query': diem_danh_query,'variable_values':variable_values, 'table_setting':diem_danh_table_setting,
                             'out_datas_func':diem_danh_out_datas_func,
                             'gen_row_data':None }  }),
     ('header3',{'range':['auto',Begin_icol_diem_danh],'val':u'Kết thúc'}),
                     ] 

    default_fixups_style  =  xlwt.easyxf(generate_easyxf(vert = 'center',horiz = 'center'))
    wb = write_fixups(fixups_diem_danh, defaut_fixups_style = default_fixups_style )
    return wb



variable_values ={ "from": "1999-01-01", "to": "2019-10-10" }  
wb = diem_danh_report_xl(variable_values)
wb.save(r'C:\Users\tu\Desktop\New folder\abc.xls')       

print('done')


