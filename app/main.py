from asyncio.log import logger
import sys
import datetime
import logging 

# from numpy import delete
from prepair import getpath
from pretask import manage_excel as me
from pretask import get_company_name as gcn
from maintask import write_excel_cost as wec
from maintask import revenue_by_lab as rbl

from logs import logerror 


# class mainactivity:
#     def __init__(self, sys_date):
#         self.sys_date = sys_date
#     #show system datetime
#     def printing(self):
#         print(self.sys_date)
#         return(self.sys_date)
#     def call_packet():
#         pass





def validate(date_text):
    try:
        sys_date = datetime.datetime.strptime(date_text, '%Y-%m-%d')
        date_only = sys_date.date()
    except ValueError as e:
        log_status = "Fail"
        logger.error(str(e)+'status:'+log_status)
        raise ValueError("Incorrect data format, should be YYYY-MM-DD")
    else:
        logger.debug('status: complete')
        return str(date_only)
class callpackage:
    def __init__(self,sys_date_run):
        self.sys_date_run = sys_date_run

    def call_all_package(self):        
        # sys_date = validate(sys.argv[1])
        print("s")
        path = getpath.main_function()
        Excel = me.Manageexcel(self.sys_date_run,path['sPath_Master'],
                               path['sPath_Data'],path['sPath_Error'],)
        ExcelName = gcn.get_company_name(self.sys_date_run,Excel,path['sPath_Temp'])
        write_file,count_y = wec.writefile(Excel,ExcelName,sys_date,path['sPath_Data'],path['sPath_Temp'],path['sPath_Template'])
        write_file2 = rbl.write_excel_revenue(self.sys_date_run,path['sPath_Data'], path['sPath_Template'],write_file,count_y)
    

        

## select system date
if __name__=='__main__':
    logger = logerror.logsetup()
    logger.debug('Start of main program')
    print("====This is main====")
    #use system input (manual case)
    if (len(sys.argv)>1):
        sys_date = validate(sys.argv[1])
        system_run = callpackage(sys_date_run = sys_date)
        system_run.call_all_package()
        # print(sys_date)
        # # classcall = mainactivity(sys_date=sys.argv[1])
        # # classcall.printing()
        # path = getpath.main_function()
        # Excel = me.Manageexcel(sys_date,path['sPath_Master'],
        #                        path['sPath_Data'],path['sPath_Error'],)
        # ExcelName = gcn.get_company_name(sys_date,Excel,path['sPath_Temp'])
        # write_file,count_y = wec.writefile(Excel,ExcelName,sys_date,path['sPath_Data'],path['sPath_Temp'],path['sPath_Template'])
        # write_file2 = rbl.write_excel_revenue(sys_date,path['sPath_Data'], path['sPath_Template'],write_file,count_y)
    #use system date (1st day and 16st of month)
    else:
        sys_date_now = datetime.datetime.now().strftime("%Y-%m-%d")
        sys_date_now = str(sys_date_now)
        system_run = callpackage(sys_date_run = sys_date_now)
        system_run.call_all_package()
        # classcall = mainactivity(sys_date=(sys_date_now.strftime("%Y-%m-%d")))
        # classcall.printing()
        # path = getpath.main_function()
        # Excel = me.Manageexcel(sys_date_now,path['sPath_Master'],
        #                        path['sPath_Data'],path['sPath_Error'],)
        # ExcelName = gcn.get_company_name(sys_date_now,Excel,path['sPath_Temp']) 
        # write_file,count_y = wec.writefile(Excel,ExcelName,sys_date_now,path['sPath_Data'],path['sPath_Temp'],path['sPath_Template'])
        # write_file2 = rbl.write_excel_revenue(sys_date_now,path['sPath_Data'],path['sPath_Template'],write_file,count_y)
        

        # print(path) #for test path value
        # print("sPath_Root : ",path['sPath_Root'])


