#libraries
from flask import Flask
from flask import render_template,request
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import re
from io import BytesIO
import os

#Configuring flask app 
app = Flask(__name__)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/",methods=["POST"])
def saveinfo():
    if request.method == "POST":
        #Retrieving data from form
        uploaded_file = request.files["excel"]
        #exported_file = request.files["export"]
        q = request.form.get('query')

        #Data Preprocessing
        excel_file = pd.read_excel(uploaded_file)
        data_frame = pd.DataFrame(excel_file)
        data_frame.drop(['Allocated_Amount__USD_','Consumed_2014','Consumed_PSWO','Allocated_Current','Active_in_CICO','Consumed_RWO','SOW_Active_Date','Total_Consumed_Amount','Service_Offering','Categories','Portfolio'],axis=1,inplace=True)
        query = re.split('[,/]',q)

        #Filtering last 3 months data
        starting_date = pd.Timestamp.now() - relativedelta(months = 3)
        data = data_frame.loc[data_frame['AS_DM'].isin(query)]
        data = data[data['Start_Date'] >= starting_date]
        
        data_col = data.columns.values.tolist()
        result = data.values.tolist()
        return render_template('index.html',result = result,data_col = data_col)

#Function to export the selected data
@app.route("/actionData",methods=["POST"])
def actionData():
    #Retrieving data
    data = request.get_data()
    d = BytesIO(data)

    #Converting bytes data to csv file
    df = pd.read_csv(d)
    items = [i for i in df.columns]
    def split(items):
        for i in range(0, len(items), 6):
            yield items[i:i + 6]
    df_data = list(split(items))

    #Converting csv file to dataframe
    col = ['SOW_','Project_Name','Start_Date','End_Date','Total SOW cost','Rev: Q1-22 (Apr-Jun2022)','Rev: Q2-22 (Jul-Sep2022)','Rev: Q3-22 (Oct-Dec2022)','Rev: Q4-23 (Jan-Mar2023)','Contract Type','Current stage','Onsite FTE','Offshore FTE','PM FTE','PM Name','TS DM name','Date SOW approved','SOW Scope Items (Summary)','Challenges','Project work status']
    df = pd.DataFrame(df_data,columns=['ind','SOW_','Project_Name','Start_Date','End_Date','AS_DM'])
    df.drop(['ind','AS_DM'],axis=1,inplace=True)
    df1 = pd.DataFrame(columns=col)
    df1['SOW_'] = df['SOW_']
    df1['Project_Name'] = df['Project_Name']
    df1['Start_Date'] = df['Start_Date']
    df1['End_Date'] = df['End_Date']
    #Checking whether the excel file exists or not
    if os.path.exists('excel_export.xlsx')!=True:
        df1.to_excel('excel_export.xlsx',index=False)
    else:
        excel_file = pd.read_excel('excel_export.xlsx')
        combined_df = pd.concat([excel_file,df1],ignore_index=True,sort=False)
        combined_df.drop_duplicates(['SOW_','Project_Name'],inplace=True,ignore_index=True)
        combined_df.to_excel('excel_export.xlsx',index=False)
    
    return "success"

if __name__ == "__main__":
    app.run(debug="True")