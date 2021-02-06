# check missing names
import pandas as pd
import numpy as np
import yaml

excel_path="C:\\Users\\encylin\\OneDrive - Ericsson AB\\Documents\\Downloads"
arg=['Jan']
project='BBSM SM6705 B261 NODE'

def wrap_list(lst, items_per_line=5):
    lines = []
    for i in range(0, len(lst), items_per_line):
        chunk = lst[i:i + items_per_line]
        line = ", ".join("{!r}".format(x) for x in chunk)
        lines.append(line)
    return "[" + ",\n ".join(lines) + "]"

def remove_duplicates_in_list(prelist):
    postlist=[]
    for item in prelist:
        if item not in postlist:
            postlist.append(item)
    return postlist

def check_missing_names(df_time,project,df_name):
    names_missing=[]
    names_recorded=df_name['Name'].tolist()
    names_to_be_checked=df_time.loc[df_time[time_project_column]==project,time_name_column].tolist()
    for item in names_to_be_checked:
        if item !='#' and item not in names_recorded:
            names_missing.append(item)
    return names_missing

def unit_month_req(u6,u7,month):
    req_u7=0
    req_u6="{:.0f}".format(df_replir.loc[(df_replir[replir_content_column] == 'Req.') & \
                                   (df_replir[replir_denomination_column] == 'Hours') & \
                                   (df_replir[replir_unit_level6_column] == u6) & \
                                   (df_replir[replir_month_column] == '2020M12') \
        , replir_value_column].sum())
    if u7 != '':
        req_u7 = "{:.0f}".format(df_replir.loc[(df_replir[replir_content_column] == 'Req.') & \
                                               (df_replir[replir_denomination_column] == 'Hours') & \
                                               (df_replir[replir_unit_level7_column] == u7) & \
                                               (df_replir[replir_month_column] == '2020M12') \
            , replir_value_column].sum())
    return [req_u6,req_u7]


def unit_month_actual(u6,u7,month,project):
    actual_u7=0
    namesinu6 = df_name.loc[df_name['Req Unit L06'] == u6, 'Name'].values.tolist()
    actual_u6="{:.0f}".format(df_time.loc[(df_time[time_name_column] != '') \
                        & (df_time[time_project_column] == project) \
                         & (df_time[time_name_column].isin(namesinu6)) , time_month_column].sum())
    if u7!='':
        namesinu7 = df_name.loc[df_name['Req Unit L07'] == u7, 'Name'].values.tolist()
        actual_u7 = "{:.0f}".format(df_time.loc[(df_time[time_name_column] != '') \
                                                & (df_time[time_project_column] == project)  \
                                                & (df_time[time_name_column].isin(namesinu7)) \
                                              , time_month_column].sum())
    return [actual_u6,actual_u7]

if __name__=='__main__':
    # get paramaters
    config=excel_path+'/'+'time_report.yaml'
    with open(config) as file:
        parameterlist = yaml.load(file, Loader=yaml.FullLoader)

    # replir file to dataframe
    replir=excel_path+'/'+parameterlist['replir']['file_name']
    df_replir = pd.read_excel(replir, sheet_name='Excel Format')
    replir_content_column=parameterlist['replir']['column_map']['Content'].rstrip()
    replir_denomination_column = parameterlist['replir']['column_map']['Denomination'].rstrip()
    replir_unit_level6_column = parameterlist['replir']['column_map']['Req Unit L06'].rstrip()
    replir_unit_level7_column = parameterlist['replir']['column_map']['Req Unit L07'].rstrip()
    replir_month_column = parameterlist['replir']['column_map']['Month'].rstrip()
    replir_value_column = parameterlist['replir']['column_map']['Value'].rstrip()
    replir_project_column = parameterlist['replir']['column_map']['Proj Name'].rstrip()

  # time report  to dataframe
    time_report=excel_path+'/'+parameterlist['timereport']['file_name_prefix']+arg[0]+parameterlist['timereport']['file_name_suffix']
    df_time = pd.read_excel(time_report)
    df_time.replace(np.nan, 0)
    time_name_column = parameterlist['timereport']['column_map']['name'].rstrip()
    time_project_column=parameterlist['timereport']['column_map']['project'].rstrip()
    time_costcenter_column = parameterlist['timereport']['column_map']['cost center'].rstrip()
    time_month_column = parameterlist['timereport']['column_map'][arg[0]].rstrip()

    # name & org  to dataframe
    organization = excel_path + '/organization.xlsx'
    df_name = pd.read_excel(organization, sheet_name='staff')
    df_unit = pd.read_excel(organization, sheet_name='unit')

    #actual report
    actual_report = excel_path + '/' + parameterlist['outputs']['actual_prefix'] + arg[0] +'_' +project+'.xlsx'
    actual_vs_plan = excel_path + '/' + arg[0] +'_' + 'output.xlsx'


    print (df_replir['Unnamed: 1'].value_counts().idxmax())



    # if check_missing_names(df_time,project,df_name)!=[]:   # to make sure namelist is completed
    #     missing_list=check_missing_names(df_time,project,df_name)
    #     print ('%s project is missing those names: ' %project)
    #     print (wrap_list(missing_list))
    #     print (' you need to add those names to the namelist then to run the acual report!')
    # else:
    #     print(' no missing names!')
    #     print(' create actual time reports!')
    #     print (df_time.columns)
    #     df_actual_report = df_time.loc[(df_time[time_name_column] != '')  & (df_time[time_project_column] == project) \
    #         , [time_costcenter_column, time_name_column, time_month_column]]
    #     df_actual_report.columns = ['cost center', 'Name', 'hours']
    #     df_actual_report.to_excel(actual_report , index=False)
    #     print(' create actual vs plan report!')
    #     data = {}
    #     data['unit'] = []
    #     data['plan'] = []
    #     data['actual'] = []
    #     for u6,u7 in df_unit.values.tolist():
    #         data['unit'] = data['unit'] + [u6, u7]
    #         data['plan'] = data['plan'] + unit_month_req(u6, u7, arg[0])
    #         data['actual'] = data['actual'] + unit_month_actual(u6, u7, arg[0], project)
    #
    #     df_output = pd.DataFrame(data, columns=['unit', 'plan', 'actual'])
    #     # print(df_output)
    #     df_output.to_excel(actual_vs_plan, index=False)

    #    print (check_missing_names(df_time,df_name,time_column_name))

#    pre_list = df_replir.loc[df_replir[project_name].str.contains("6705", na=False), project_name].values.tolist()
 #   post_list=remove_duplicates_in_list(pre_list)
  #  print (post_list)

