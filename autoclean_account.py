import pandas
import os
import datetime
import numpy as np
import time

def f_auto_clear (gl_account, company_code, clear_date):
    str_h='''
    If Not IsObject(application) Then
       Set SapGuiAuto  = GetObject("SAPGUI")
       Set application = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(connection) Then
       Set connection = application.Children(0)
    End If
    If Not IsObject(session) Then
       Set session    = connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session,     "on"
       WScript.ConnectObject application, "on"
    End If
    session.findById("wnd[0]").resizeWorkingPane 100,23,false
    session.findById("wnd[0]/tbar[0]/okcd").text = "f-03"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/sub:SAPMF05A:0131/radRF05A-XPOS1[4,0]").select
    
    '''
    f1 = open ('autoclear_105990.vbs','w')
    f1.write (str_h)
    f1.write ('session.findById("wnd[0]/usr/ctxtRF05A-AGKON").text = "' + gl_account + '"'   + '\n' )
    f1.write ('session.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text = "' + clear_date +  '"' + '\n')
    f1.write ('session.findById("wnd[0]/usr/txtBKPF-MONAT").text = "' +  clear_date[0:2]   +   '"' + '\n')
    f1.write ('session.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = "'+ company_code +'"' + '\n')
    f1.write ('session.findById("wnd[0]/usr/sub:SAPMF05A:0131/radRF05A-XPOS1[4,0]").setFocus' + '\n')
    f1.write ('session.findById("wnd[0]/tbar[0]/btn[11]").press' + '\n')
    f1.write ('session.findById("wnd[0]/usr/sub:SAPMF05A:0732/ctxtRF05A-VONDT[0,0]").text = "'  + clear_date +  '"' + '\n')
    f1.write ('session.findById("wnd[0]/usr/sub:SAPMF05A:0732/ctxtRF05A-VONDT[0,0]").caretPosition = 10' + '\n')
    f1.write ('session.findById("wnd[0]/tbar[1]/btn[16]").press' + '\n')
    f1.write ('session.findById("wnd[0]/tbar[0]/btn[11]").press' + '\n')
    f1.close()
    os.system('autoclear_105990.vbs')




str_header='''

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").resizeWorkingPane 100,23,false
session.findById("wnd[0]/tbar[0]/okcd").text = "/nfbl3n"
session.findById("wnd[0]").sendVKey 0


'''

str_tail='''

session.findById("wnd[0]/usr/txtPA_NMAX").setFocus
session.findById("wnd[0]/usr/txtPA_NMAX").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press


'''
str_account = raw_input ('GL Account [105990]:') or '105990'

str_now_date = datetime.datetime.now().strftime('%m/%d/%Y')
str_date = raw_input ('AS OF DATE [' + str_now_date + ']' ) or str_now_date


f2 = open ('download_105990.vbs','w')
f2.write (str_header)
f2.write ('session.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").text = "'+  str_account +'"' +'\n')
f2.write ('session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text = "8002"' +'\n')
f2.write ('session.findById("wnd[0]/usr/ctxtSD_BUKRS-HIGH").text = "8003"' +'\n')
f2.write ('session.findById("wnd[0]/usr/ctxtPA_STIDA").text = "' + str_date  +  '"' +'\n' )
f2.write ('session.findById("wnd[0]/usr/ctxtPA_VARI").text = "MAX-105990s1"' +'\n')
f2.write (str_tail)
f2.close()

print ('SAP Processing... Please Wait...')
os.system('download_105990.vbs')


data=pandas.read_clipboard(sep='|', skiprows=6 )
data= data[data.CoCd.isin(['8002','8003'])]
del data['Unnamed: 0']
del data['Unnamed: 4']
data.columns=['POSTINGDATE','COMPANYCODE','AMOUNT']



data['POSTINGDATE']=data['POSTINGDATE'].map(str.strip)
data['COMPANYCODE']=data['COMPANYCODE'].map(str.strip)
data['AMOUNT']=data['AMOUNT'].astype('str')
data['AMOUNT']=data['AMOUNT'].map(str.strip)


for row in data.iterrows():
    idx=row[0]-1
    #print data.values[idx][0]
    data.values[idx][0]=data.values[idx][0][2:12]

data2= data[data.AMOUNT.isin(['0.00','0.0'])]
print ('SAP Data acquired.')
print data2

idx=0
for row in data2.iterrows():    
    str_post_date=data2.values[idx][0]
    str_company_code=data2.values[idx][1]
    print ("Processing Row " + str(idx+1) + ' of ' + str(data2.shape[0]) )
    f_auto_clear(str_account, str_company_code, str_post_date)
    if idx+1<data2.shape[0]:
        time.sleep(15)
    idx=idx+1
    





