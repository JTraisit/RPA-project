from asyncio.log import logger
import openpyxl as xl
import os
import sys

from tkinter import Variable
from logs import logerror

#getpath
def main_function(): # main function for calling another function.
    logger = logerror.logsetup()
    print("getpath processing...\n")
    excel_path = get_excel_path() # call get_excel_path function to get a path file
    path = read_data(rawdata_path=excel_path,logger=logger) # read data to get path in excel file and store in dictionary.
    return path


def get_excel_path():
    keep = sys.path[0] # get path file by sys library.
    excel_path = (str(keep)+'\master\Path.xlsx') # adding path master and file Path.xlsx 
    return excel_path


def read_data(rawdata_path,logger):
    try:
        file =  xl.load_workbook(rawdata_path,data_only=True) # read excel to get data (data_only to get a value of excel formula)
    except FileNotFoundError as f:
        log_status = "Fail"
        logger.error(str(f)+'status:'+log_status)
        raise FileNotFoundError("File path in Master folder not found, please check your file and try again later.")

    
    else:
        logger.debug('status: complete')
        path = {}
        path_name = ["sPath_Root","sPath_Master","sPath_Data","sPath_Error","sPath_Temp","sPath_Template"]
        position = ['B2','B3','B4','B5','B7','B8']
        file.active = file['Path']
        if (len(path_name)==len(position)):
            for i in range(len(path_name)):
                path[path_name[i]] = file.active[position[i]].value
        else:
            print("Error number of path_name and position are not equal please check and try again later")
    # try:
    #     file =  xl.load_workbook(rawdata_path,data_only=True) # read excel to get data (data_only to get a value of excel formula)
    # except FileNotFoundError:
    #     print("File path in Master folder not found, please check your file and try again later.")
    #     raise


    #     # get value in each call.
    #     sPath_Root = file.active['B2'].value
    #     sPath_Master = file.active['B3'].value
    #     sPath_Data = file.active['B4'].value
    #     sPath_Error = file.active['B5'].value
    #     sPath_Temp = file.active['B7'].value
    #     sPath_Template = file.active['B8'].value
    # # Create dictionary path
    #     path = {
    #      "sPath_Root":sPath_Root,
    #      "sPath_Master":sPath_Master,
    #      "sPath_Data":sPath_Data,
    #      "sPath_Error":sPath_Error,
    #      "sPath_Temp":sPath_Temp,
    #      "sPath_Template":sPath_Template
    #     }
        print("getpath success! return:"+str(path)+"\n") 
        return path
  
   



