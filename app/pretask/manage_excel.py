#manage excel
import pandas as pd


def Manageexcel(sys_date,sPath_Master,sPath_Data,sPath_Error):
    print("manage excel  processing...\n")
    sys_date_split = sys_date.split("-")
    sys_day = sys_date_split[2]
    sys_month = sys_date_split[1]
    sys_year = sys_date_split[0]


    try:
        def read_data(sPath_Master):
            df = pd.read_excel((sPath_Master+'/Raw_Data.xlsx'),sheet_name=0)
            return df
    
        def display(df):
            # pd.set_option('display.max_columns', None) #set ค่าแสดงผลทั้งหมด แต่ต้องใช้สำหรับ df.head()
            # pd.set_option('display.max_rows', None)
            print(df.head())
        
        def select(df):
            return df[["moddt",
                            "s_requestid",
                            "quotation",
                            "u_company",
                            "requestdt",
                            "u_forlab"]]
        
        def drop_na(df):
            return (df.dropna())
    
        def select_between_period(df):
            df = df.set_index(['requestdt'])
            # 1st run
            if(int(sys_day) == 1):
                #[run on (1/1/xxxx) select {1/12/(xxxx-1) - 31/12/(xxxx-1)}]
                if(int(sys_month) == 1):
                    sel_sys_month = 12
                    sel_sys_year = (int(sys_year)-1)
                #[run on (1/2/xxxx)-(1/12/xxxx)]
                else:
                    if (len(str(int(sys_month)-1)) == 1): # when the month(RUN ON) is the unit [02-10] => [1-9]
                        sel_sys_month = str('0'+(str(int(sys_month)-1))) #[1-9] => [01-09]
                    else:
                        sel_sys_month = str((str(int(sys_month)-1)))#[10-12]
                    sel_sys_year = sys_year
                sel_period = (str(sel_sys_year)+'-'+str(sel_sys_month))
                df_sel_period = df.loc[sel_period]  
                df_sel_period = df_sel_period.rename_axis(['requestdt']).reset_index()  
            elif((int(sys_day)>1)&(int(sys_day)<= 31)):
                sel_period_L = ((str(sys_year)+'-'+str(sys_month)+'-'+'01'))
                sel_period_R = (str(sys_year)+'-'+str(sys_month)+'-'+str(int(sys_day)-1))
                df_sel_period = df.loc[sel_period_L:sel_period_R]   
                df_sel_period = df_sel_period.rename_axis(['requestdt']).reset_index() 
            else:
                #errorcase
                print("System date has problem, please check and try again.")
            
            return df_sel_period
    
        def group_by_sel_maxmoddt(df):
            gb1 = df.groupby(['s_requestid','quotation','u_company'], sort=False,as_index = False)['moddt'].max()
            gb2 = df.groupby(['s_requestid','u_forlab'], sort=False,as_index = False)['moddt'].max()
            combine = pd.merge(gb1, gb2, how ='inner', on =['s_requestid', 'moddt']) # inner join gb1 & gb2
            df_groupby = combine.groupby(['u_company','u_forlab'], as_index = False)['quotation'].sum() #groupby u_company and u_forlab by sum quatation
            
            return df_groupby
    
        def write_file(df):
            df.to_excel(sPath_Data+'/Input/'+sys_date+'.xlsx', index=True)
            # df.to_excel(sPath_Data+'/Input/'+'2022-01-01--2022-01-31'+'.xlsx', index=True)
            
    
                
            
        #call function
        # display(keep1) # for display dataframe
        keep = read_data(sPath_Master) #step 1 read data.
        keep = select(keep)  #step 2 select 6 column.
        keep = drop_na(keep) #step 3 drop empty value.
        keep = select_between_period(keep) #step 4 select data between period from the condition based on system up-time.
        output = group_by_sel_maxmoddt(keep) #step 5 groupby data to obtain the smallest unit of data.
        write_file(output) #step 6 write an excel file with xlsx file extension.
        # display(output)
        print("manage excel success!!\n") 
    
        return output
    except FileNotFoundError:
        raise FileNotFoundError("File Raw Data in Master folder not found, please check your file and try again later.")
        
        
        #cache file
    
        # if(((int(sys_month) >= 1) & (int(sys_month)<=12) &( int(sys_day) == 1))):
        #     output.to_excel(sPath_Temp+'Cache-'+sys_date+'.xlsx', index=True)
    
        # gb = keep.groupby(['s_requestid','quotation'], sort=False)['moddt'].max()
        # print(gb)

    
   
    

  




