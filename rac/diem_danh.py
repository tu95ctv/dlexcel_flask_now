# -*- coding: utf-8 -*-
import xlwt
# from datetime import datetime
from dltool import generate_easyxf as generate_easyxf_import
from dltool import write_table_rerange, common_one_table_report_xl, font_decorator_parent_new, convert_gmt_str_dt_to_vn_str_dt, easyxf_new
from dltool import display_from_to

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
        datas_diem_danh =rs['data']['result']
        print ('**datas_diem_danh', datas_diem_danh)
        return datas_diem_danh


    @font_decorator_parent_new(font = font, height = font_size )
    def generate_easyxf(*args, **kargs):
        return generate_easyxf_import(*args, **kargs)
    
    diem_danh_FIELDNAME_FIELDATTR = [
            ('stt',{'val_func': lambda v,d,s:s['stt'].get('val',0)+1,'width':6, 'title':'STT'}),
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
            'title_height':20*26
            }
    return diem_danh_table_setting

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
if __name__ == '__main__':
    variable_values_dd ={ "from": "1999-01-01", "to": "2019-10-10", 'break_sheet':True }  
    wb = diem_danh_report_xl(variable_values_dd)
    wb.save(r'C:\Users\tu\Desktop\New folder\diem_danh.xls')       
    print('done')


