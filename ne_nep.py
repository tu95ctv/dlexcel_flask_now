# -*- coding: utf-8 -*-
import xlwt
from datetime import datetime
from dltool import generate_easyxf as generate_easyxf_import
from dltool import font_decorator_parent, write_hasura_table,write_fixups
from dltool import display_from_to
# Font = 'Calibri'


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
    
def ne_nep_report_xl(variable_values,  font_size = None, font = 1):

    
#     font = font_map.get(font,None)
    @font_decorator_parent(font, font_size)
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
    table_setting = {
            'default_width':13,
            'default_merge_title_font':xlwt.easyxf(generate_easyxf(bold=True,vert = 'center',horiz='center',borders='left thin, right thin, top thin, bottom thin')),
            'default_title_font':xlwt.easyxf(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center',horiz = 'center')),
            'default_cell_font':xlwt.easyxf(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center',horiz = 'center')),
            'is_merge_title':True,
            'FIELDNAME_FIELDATTR':FIELDNAME_FIELDATTR_ne_nep
            }
    Beign_irow = 2
    Begin_icol = 0
    
    
    
    fixups =[  
         ('header',{'range':[Beign_irow, Beign_irow, Begin_icol, Begin_icol+7],'val':u'Bảng nề nếp', 'style':xlwt.easyxf(generate_easyxf(bold=True,height=20, vert = 'center',horiz = 'center'))}),
         ('header2',{'range':['auto', 0, Beign_irow, Begin_icol+7],'val_func':display_from_to, 'val_func_kargs':{'variable_values':variable_values}, 'style':xlwt.easyxf(generate_easyxf( vert = 'center',horiz = 'center'))}),
         ('table',{'range':['auto', Begin_icol],'val':None,
                   'func':write_hasura_table, 'offset':2,
                   'func_kargs':{'query': ne_nep_query,'variable_values':variable_values, 'table_setting':table_setting,
                                 'out_datas_func':ne_nep_out_datas_func,
                                 'gen_row_data':ne_nep_gen_row_data }  }),
         ('header3',{'range':['auto',Begin_icol],'val':u'Kết thúc'}),
                     ] 

    default_fixups_style  =  xlwt.easyxf(generate_easyxf(vert = 'center',horiz = 'center'))
    wb = write_fixups(fixups, defaut_fixups_style = default_fixups_style )
    return wb



variable_values ={ "from": "1999-01-01", "to": "2019-10-10" }  
wb = ne_nep_report_xl(variable_values)
wb.save(r'C:\Users\tu\Desktop\New folder\abc.xls')       

print('done')


