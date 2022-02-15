# Package import
import datetime
import os

from flask import Flask, render_template, send_file, make_response, url_for, Response, redirect, request, \
    send_from_directory
import pandas as pd
import openpyxl

# initialise app
app = Flask(__name__)


def out(d1, d2, d3):
    Tag = "Shared"  ## Has to be Onshore/Shared/Offshore
    import numpy as np
    import pandas as pd
    import datetime
    import calendar
    pd.set_option('display.max_rows', 500)
    pd.set_option('display.max_columns', 500)
    pd.set_option('max_colwidth', 100)
    import calendar
    month_name = int((datetime.datetime.now() - pd.DateOffset(months=1)).strftime('%m'))
    month_name = calendar.month_abbr[month_name]
    month_name
    # #### Mention path for the SF Oracle Distribution ID file.
    ### TODO: SF_Oracle file from finance team in the zip folder
    oap = pd.read_excel(d1)
    oap = oap[~pd.isna(oap['Oracle AP Distribution ID'])]
    oap = oap[oap['Oracle AP Distribution ID'].str.len() == 52]
    oap['Service Line'] = oap['Oracle AP Distribution ID'].apply(lambda x: str(x).split('.')[4])
    oap['Program'] = oap['Oracle AP Distribution ID'].apply(lambda x: str(x).split('.')[5])
    oap['Project'] = oap['Oracle AP Distribution ID'].apply(lambda x: str(x).split('.')[6])
    oap = oap[['Project Name', 'Project', 'Program', 'Service Line']]
    oap = oap.drop_duplicates()
    oap.head()
    # ### Allocating cost for available teams
    # #### Add file path for project split up sheet
    # dfl=pd.read_excel(f"Input/Onshore/InputCost_Onshore_{month_name}21/Project_onshore_cost_{month_name}-21.xlsx",None)
    # dfl=pd.read_excel(f"Input/Shared/SharedMRR_Input_{month_name}2021/Shared_Cost_{month_name}-21_V1.xlsx",None)
    dfl = pd.read_excel(d2, None)
    dfl1 = []
    type1 = []
    for type_, df in dfl.items():
        print(type_)
        type1.append(type_)
        df['Cost Type'] = type_
        dfl1.append(df)
    split = pd.concat(dfl1, sort=False)
    split['% Cost'] = split['% Cost'] * 100
    split.head()
    if Tag == "Onshore":
        test_list = ['Onshore MCC', 'HRA Clinical', 'HRA Coding', 'HRA Outreach', 'HRA General',
                     'HRA Scheduling Coordinators']
    elif Tag == "Shared":
        test_list = ['MRT cost', 'MRR Back Office', 'OC cost']
    else:
        test_list = ['Annova', 'Visaya']
    set(type1)
    set(test_list)
    if set(type1) != set(test_list):
        raise ValueError
    split.dtypes
    split.columns
    split.groupby('Cost Type')['% Cost'].sum()
    if round(split.groupby('Cost Type').sum()['% Cost'].sum(), 2) != float(len(type1) * 100):
        raise ValueError
    conv = {'Account': str, 'Branch': str, 'Cost Centre': str, 'Service Line': str, 'Program': str, 'Project': str,
            'Future1': str}
    dfl = pd.read_excel(d3, None, converters=conv)
    dfl1 = []
    type2 = []
    for type_, df in dfl.items():
        print(type_)
        type2.append(type_)
        df['Cost Type'] = type_
        dfl1.append(df)
    cost = pd.concat(dfl1)
    cost.head()
    if set(type2) != set(type1):
        raise ValueError
    cost.groupby('Cost Type')["Total"].sum()
    fnl = cost.merge(split, on='Cost Type', how='left')
    fnl['Entered DR'] = ((fnl['Total'] * fnl['% Cost']) / 100).round(2)
    fnl['Entered CR'] = 0
    fnl.drop(['Program', 'Project', 'Service Line', '% Cost', 'Total'], axis=1, inplace=True)
    fnl
    fnl[pd.isna(fnl['Project Name'])]  ## TODO: This should be empty, error contact Shivakanth
    if fnl[pd.isna(fnl['Project Name'])].shape[0] != 0:
        raise ValueError
    # ### Everything should be zero in the below cell
    fnl.isna().sum()  ## Only Visit type column in onshore sheet will have non zero value here.
    # fnl[fnl['Cost Type']=='MRT cost with Konnect']
    # ### Finding Blank Oracle AP id
    a = fnl['Project Name'].unique()
    b = oap['Project Name'].unique()
    c = [i for i in a if i not in b]
    c  ## This should be blank
    if len(c) > 0:
        raise ValueError
    # ### This should be blank
    fnl1 = fnl.merge(oap, on='Project Name', how='left')
    fnl1[pd.isna(fnl1['Project'])]['Project Name'].unique()
    fnl1.groupby(["Cost Type"])["Entered DR"].sum()
    fnl2 = fnl1.drop(['Project Name'], axis=1)
    fnl2.head()
    # fnl2.drop(labels=['Unnamed: 3','Unnamed: 4'],axis = 1,inplace=True)
    # ### Reversing the Credit amount : Adding previous debit to Credit
    # reverse.groupby("Cost Type")["Entered CR"].sum()
    cost.head()
    reverse = cost.copy()
    reverse['Entered DR'] = 0
    reverse['Entered CR'] = reverse['Total']
    reverse['Project Type'] = 'NA_Reverse_Entry'
    if Tag == 'Onshore':
        reverse['Visit Type'] = np.nan
    reverse.drop('Total', axis=1, inplace=True)
    reverse = reverse[fnl2.columns]
    reverse
    final = pd.concat([fnl2, reverse])
    final.head()
    # ### Final Output Generation
    ## TODO: Add mapping incase any new column is added.
    cl1 = ['Account', 'Branch', 'Cost Centre', 'Service Line', 'Program', 'Project', 'Future1']
    cl2 = ['Segment2', 'Segment3', 'Segment4', 'Segment5', 'Segment6', 'Segment7', 'Segment9']
    cold = dict(zip(cl1, cl2))
    cold
    if len([i for i in cl1 if i not in final.columns]):
        print("ERROR:   Column name absent", [i for i in cl1 if i not in final.columns])
    final.columns
    c4 = final.rename(columns=lambda x: cold.get(x, x))
    c4
    today = datetime.date.today()
    str(today)
    first = today.replace(day=1)
    lastMonth = first - datetime.timedelta(days=1)
    str(lastMonth)
    month_name = calendar.month_abbr[lastMonth.month]
    year = lastMonth.year
    year = str(year)[2:]
    year, month_name  ## These variables will used to name the output file.
    # Dec = lastMonth - datetime.timedelta(days=31)
    # Dec
    oracle = pd.read_excel('output_v1.xlsx')
    oracle.head(1)
    # or_map= pd.read_excel('output_v1.xlsx','map')
    # or_map
    c4['*Journal Category'] = 'Manual'
    c4['*Currency Code'] = 'USD'
    c4['*Journal Entry Creation Date'] = str(today)
    c4['*Actual Flag'] = 'A'
    c4['Segment1'] = '11'
    c4['Segment8'] = '00'
    # c4['Segment9']='0000'
    c4['Segment10'] = '0000'
    c4['*Status Code'] = 'NEW'
    c4['*Ledger ID'] = '300000005557718'
    c4['*Effective Date of Transaction'] = str(lastMonth)
    c4['*Journal Source'] = 'Manual'
    scll = c4.columns
    fcll = oracle.columns
    scll
    if Tag == 'Onshore':
        fcll = ['Cost Type', 'Project Type', 'Visit Type'] + list(fcll)
    else:
        fcll = ['Cost Type', 'Project Type'] + list(fcll)
    if len([i for i in scll if i not in fcll]) > 0:
        raise ValueError
    ### should come 97 23 74 97 for Onshore else 96 22 74 96
    a_cll = [i for i in fcll if i not in scll]
    print(len(fcll), len(scll), len(a_cll), len(scll) + len(a_cll))
    scll
    for col in a_cll:
        c4[col] = ""
    c4 = c4[fcll]
    c4.head()
    diff = c4.groupby('Cost Type', as_index=False)['Entered DR', 'Entered CR'].sum()
    diff['dif'] = diff['Entered CR'] - diff['Entered DR']
    diff['dif'] = diff['dif'].round(2)
    diff
    dif_dict = dict(zip(diff['Cost Type'], diff['dif']))
    dif_dict
    tdfl = []
    for ct in c4['Cost Type'].unique():
        tdf = c4.loc[c4['Cost Type'] == ct, :]
        tdf.reset_index(drop=True, inplace=True)
        tdf.loc[0, 'Entered DR'] = tdf.loc[0, 'Entered DR'] + dif_dict[ct]
        tdfl.append(tdf)
    c5 = pd.concat(tdfl)
    c5.head()
    c4.shape, c5.shape
    c4[pd.isna(c4['Segment5'])]['Entered DR'].sum()
    ## Should come as False
    c4['Entered CR'].sum(), c4['Entered DR'].sum(), round(c4['Entered CR'].sum(), 2) == c4['Entered DR'].sum()
    if round(c4['Entered CR'].sum(), 2) == c4['Entered DR'].sum() == True:
        raise ValueError
    ## Should come as True
    c5['Entered CR'].sum(), c5['Entered DR'].sum(), round(c5['Entered CR'].sum(), 2) == round(c5['Entered DR'].sum(), 2)
    if round(c5['Entered CR'].sum(), 2) == round(c5['Entered DR'].sum(), 2) == False:
        raise ValueError
    diff1 = c5.groupby('Cost Type', as_index=False)['Entered DR', 'Entered CR'].sum()
    diff1['dif'] = diff1['Entered CR'] - diff1['Entered DR']
    diff1['dif'] = diff1['dif'].round(2)
    diff1
    c5.shape
    oracle.shape
    if Tag == 'Onshore' and c5.drop(['Visit Type'], axis=1).isna().sum().sum() != 0:
        raise ValueError
    elif Tag != 'Onshore' and c5.isna().sum().sum() != 0:
        raise ValueError
    c5.isnull().sum()
    return c5


global gpath


# decorator for homepage
@app.route('/')
def index():
    return render_template('index.html',
                           PageTitle="Landing page")


# These functions will run when POST method is used.
@app.route('/', methods=["POST"])
def return_out():
    # gathering file from form
    uploaded_file1 = request.files['file1']
    uploaded_file2 = request.files['file2']
    uploaded_file3 = request.files['file3']

    # making sure its not empty
    if uploaded_file1.filename != '' and uploaded_file2.filename != '' and uploaded_file3.filename != '':
        # d1 = pd.read_csv(uploaded_file1)
        # d2 = pd.read_csv(uploaded_file2)
        # d3 = pd.read_csv(uploaded_file3)
        # csv = pd.concat([d1, d2, d3]).to_csv()
        csv = out(uploaded_file1, uploaded_file2, uploaded_file3)
        name = uploaded_file1.filename + uploaded_file2.filename + uploaded_file3.filename
        global gpath
        gpath = 'output/' + name + ".csv"
        csv.to_csv(gpath, index=False)
        # print("sending file...")
        # result = send_file(path, as_attachment=True)
        # print("file sent, deleting...")
        # file=os.remove(path)
        # print(gpath)
        return render_template('download.html', PageTitle="Download Page")
        # return response
    else:
        return render_template('index.html',
                               PageTitle="Landing page")
    # This just reloads the page if no file is selected and the user tries to POST.


@app.route('/return-files/')
def return_files_tut():
    try:
        return send_file(gpath, as_attachment=True)
    except Exception as e:
        return str(e)


if __name__ == '__main__':
    app.run(debug=True, port=5000)
