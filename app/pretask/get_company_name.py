#get_company_name
import pandas as pd
# import manage_excel as me
import os



def get_company_name(sys_date,dataframe,sPath_Temp):
    print("get company name processing...\n")
    sys_date_split = sys_date.split("-")
    sys_day = sys_date_split[2]
    sys_month = sys_date_split[1]
    sys_year = sys_date_split[0]
    cache_file_status = False

    def sel_comp_drop_duplicate(df):# select 'u_company' column and drop duplicates row. after that, get output into dataframe and return.
        df = df[['u_company']] # select 'u_company' column.
        df = df.drop_duplicates(keep='first').reset_index(drop = True).rename_axis(['index'])
        return df
    def combind_row(df1,df2):# merge 2 table by bindrow 2 dataframe when dataframe 1 is main. after that, drop duplicates row and return. (dataframe)
        datacombind = pd.concat([df1,df2]) #bindrow
        datacombind = datacombind.drop_duplicates(keep='first').reset_index(drop = True).rename_axis(['index']) # drop duplicates row and reset index 
        return datacombind
    def write_file(df): # write cache file.
        df.to_excel(sPath_Temp+'/Cache-'+sys_date+'.xlsx', index=True)
    def delete_file(file_path):# delete cache file
        if os.path.exists(file_path): #check file is existing
            os.remove(file_path)
        else:
            print("The file does not exist") # return error
    def createdataframe(inputexcel1,dataframe):
        try: 
            dataframe1 = pd.read_excel(inputexcel1,sheet_name=0)# read excel data(month - 1)
            # delete_file(inputexcel1) #remove cache file
            dataframe1 = sel_comp_drop_duplicate(dataframe1) #use a select and drop duplicate function (month - 1)
            # dataframe1.to_excel(sPath_Temp+'/CheckDF1.xlsx', index=True)
            dataframe2 = sel_comp_drop_duplicate(dataframe)  #use a select and drop duplicate function month
            # dataframe2.to_excel(sPath_Temp+'/CheckDF2.xlsx', index=True)
            output = combind_row(dataframe1,dataframe2)
            if(sys_day == 1):
                cache_file_status = True # assign status True for cache file to temp folder.
            return output
        except FileNotFoundError:
            raise FileNotFoundError("File cache in temp folder not found, please check your file and try again later.")


    #call function
    if (((int(sys_month) == 1)&(int(sys_day) == 1))|(int(sys_month) > 2)): #when month = 1 or morethan 2 (not equal 2)
        print("asd")

        if (int(sys_day) == 1):
           if (len(str(int(sys_month)-1)) == 1): #when month - 1 are the unit ex.[02-10] => [1-9]
               inputexcel1 = (sPath_Temp+'/cache-'+(sys_year+'-0'+(str(int(sys_month)-1))+'-'+'01')+'.xlsx')  #[1-9] => [01-09]
           elif((int(sys_month) == 1)):#when system run-on month equal 1
               inputexcel1 = (sPath_Temp+'/cache-'+(str(int(sys_year)-1)+'11'+'-'+'01')+'.xlsx') #read cache file on november last month (December is a datafram) 
           else:
               inputexcel1 = (sPath_Temp+'/cache-'+(sys_year+'-'+(str(int(sys_month)-1))+'-'+'01')+'.xlsx') #[10]read cache file month-1 ex. Run-on 1/11/2022, read cache file Run-On 1/10/2022
        elif((int(sys_day)>1)&(int(sys_day)<=31)):#adding on 8/7/2022
            if (len(str(int(sys_month)-1)) == 1): #when month - 1 are the unit ex.[02-10] => [1-9] #adding on 8/7/2022
               inputexcel1 = (sPath_Temp+'/cache-'+(sys_year+'-0'+(str(int(sys_month)))+'-'+'01')+'.xlsx')  #[1-9] => [01-09] #adding on 8/7/2022
            else:
               inputexcel1 = (sPath_Temp+'/cache-'+(sys_year+'-'+(str(int(sys_month)-1))+'-'+'01')+'.xlsx') #[10]read cache file month ex. Run-on 16/5/2022, read cache file Run-On 1/5/2022 #adding on 8/7/2022
        output = createdataframe(inputexcel1,dataframe=dataframe)
        
    elif((int(sys_day)>1)&(int(sys_day)<=31)&((int(sys_month) == 2))):
        if (len(str(int(sys_month)-1)) == 1): #when month - 1 are the unit ex.[02-10] => [1-9] #adding on 8/7/2022
            inputexcel1 = (sPath_Temp+'/cache-'+(sys_year+'-0'+(str(int(sys_month)))+'-'+'01')+'.xlsx')  #[1-9] => [01-09] #adding on 8/7/2022
        else:
            inputexcel1 = (sPath_Temp+'/cache-'+(sys_year+'-'+(str(int(sys_month)-1))+'-'+'01')+'.xlsx') #[10]read cache file month ex. Run-on 16/5/2022, read cache file Run-On 1/5/2022 #adding on 8/7/2022
        output = createdataframe(inputexcel1,dataframe=dataframe)

        
    elif ((int(sys_day) == 1)&((int(sys_month) == 2))):
        output = sel_comp_drop_duplicate(dataframe)
        cache_file_status = True

    elif ((int(sys_day)>1)&(int(sys_day)<=31)&((int(sys_month) == 1))):
        output = sel_comp_drop_duplicate(dataframe)
        cache_file_status = False

    else:
        print("Error")
        # output = sel_comp_drop_duplicate(dataframe)
        # cache_file_status = False # assign status False for NOT cache file.
    if(cache_file_status): #check  cache file status
        write_file(output) #if cache file status is True use write_file function by passing output argument.
    print("get company name success!!!\n") 
    return output




