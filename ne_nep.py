# -*- coding: utf-8 -*-
import xlwt
from dltool import generate_easyxf as generate_easyxf_import
from dltool import  write_table_rerange, common_one_table_report_xl, font_decorator_parent_new
from dltool import display_from_to
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
            ('stt',{'val_func': lambda v,d,s:s['stt'].get('val',0)+1,'width':4,'title':u'STT'}),
            ('class_name',{'title':u'Lớp', 'width':5}),
            ('slg_loi_tap_the',{'merge_title':u'Vi phạm tập thể','title':u'Số lỗi vi phạm'}),
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
            'gen_row_data':ne_nep_gen_row_data,
            'default_width':11,
            'default_merge_title_font':xlwt.easyxf(generate_easyxf(align_wrap=True, bold=True,vert = 'center',horiz='center',borders='left thin, right thin, top thin, bottom thin')),
            'default_title_font':xlwt.easyxf(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center',horiz = 'center')),
            'default_cell_font':xlwt.easyxf(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center')),
            'is_merge_title':True,
            'FIELDNAME_FIELDATTR':ne_nep_FIELDNAME_FIELDATTR,
            'row_height':20*16,
            'title_height':20*26
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
if __name__ == '__main__':
    variable_values_dd ={ "from": "1999-01-01", "to": "2019-10-10", 'font_size':11 }  
    wb = ne_nep_report_xl(variable_values_dd)
    wb.save(r'C:\Users\tu\Desktop\New folder\ne_nep_tong_hop.xls')       
    print('done')


