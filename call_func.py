# -*- coding: utf-8 -*-
from ne_nep import ne_nep_report_xl
from diem_danh import diem_danh_report_xl
from nn_chi_tiet import nn_chi_tiet_report_xl
import re

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
    
if __name__ == '__main__':
    variable_values_dd ={ "from": "1999-01-01", "to": "2019-10-10", 'break_sheet': True  }  
    wb, filename = get_funcxl_and_run_funcxl_from_key('nn_chi_tiet', variable_values_dd)
    wb.save(r'C:\Users\tu\Desktop\New folder\%s'%filename)      
    
#     wb, filename = get_funcxl_and_run_funcxl_from_key('diem_danh', variable_values_dd)
#     wb.save(r'C:\Users\tu\Desktop\New folder\%s'%filename)      
#     
#     
#     wb, filename = get_funcxl_and_run_funcxl_from_key('ne_nep', variable_values_dd)
#     wb.save(r'C:\Users\tu\Desktop\New folder\%s'%filename)      
    
    print ('done')