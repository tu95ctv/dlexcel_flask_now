# -*- coding: utf-8 -*-
import xlwt
from dltool import generate_easyxf as generate_easyxf_import
from dltool import write_table_rerange, common_one_table_report_xl, font_decorator_parent_new, convert_gmt_str_dt_to_vn_str_dt, easyxf_new
from dltool import display_from_to
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

    def ne_nep_out_datas_func (rs):
        datas =rs['data']['result1']['nodes']
        return datas
    @font_decorator_parent_new(font = font, height = font_size, vert = 'center',  horiz = None )
    def generate_easyxf(*args, **kargs):
        return generate_easyxf_import(*args, **kargs)
    ne_nep_FIELDNAME_FIELDATTR = [
            ('stt',{'val_func': lambda v,d,s:s['stt'].get('val',0)+1,'width':4,'title':u'STT'}),
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
            'out_datas_func':ne_nep_out_datas_func,
            'gen_row_data':None,
            'default_width':11,
            'default_merge_title_font':xlwt.easyxf(generate_easyxf(align_wrap=True, bold=True,vert = 'center',horiz='center',borders='left thin, right thin, top thin, bottom thin')),
            'default_title_font':xlwt.easyxf(generate_easyxf(bold = True, borders='left thin, right thin, top thin, bottom thin',vert = 'center',horiz = 'center', pattern = 'pattern solid, fore_colour gray25')),
            'default_cell_font':xlwt.easyxf(generate_easyxf(borders='left thin, right thin, top thin, bottom thin',vert = 'center')),
            'is_merge_title':False,
            'FIELDNAME_FIELDATTR':ne_nep_FIELDNAME_FIELDATTR,
            'row_height':20*16,
            'title_height':20*26
            }
    return ne_nep_table_setting


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
        return datas

    @font_decorator_parent_new(font = font, height = font_size, vert = 'center',  horiz = None )
    def generate_easyxf(*args, **kargs):
        return generate_easyxf_import(*args, **kargs)
    ne_nep_tap_the_FIELDNAME_FIELDATTR = [
            ('stt',{'val_func': lambda v,d,s:s['stt'].get('val',0)+1,'width':4,'title':u'STT'}),
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
            'title_height':20*26
            }
    return table_setting
def ne_nep_gen_fixups(font, font_size,variable_values, table_setting_list, request_args):
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
    wb = common_one_table_report_xl(request_args, Basic_setting, [nn_chi_tiet_ca_nhan_gen_table_setting, nn_chi_tiet_tap_the_gen_table_setting], ne_nep_gen_fixups)
    return wb

if __name__ == '__main__':
    variable_values_dd ={ "from": "1999-01-01", "to": "2019-10-10", 'break_sheet': False  }  
    wb = nn_chi_tiet_report_xl(variable_values_dd)
    wb.save(r'C:\Users\tu\Desktop\New folder\ne_nep_chi_tiet.xls')       
    print('done')


