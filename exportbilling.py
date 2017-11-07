# coding:utf-8
import os
import datetime
import time
import pandas
str_head='''
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
session.findById("wnd[0]").resizeWorkingPane 158,34,false
session.findById("wnd[0]/tbar[0]/okcd").text = "/nvf03"
session.findById("wnd[0]/tbar[0]/btn[0]").press
'''

str_tail='''

session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[1]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press

'''
print "1. Open FBL5N in SAP."
print "2. Export result into billing_list.txt."
print "3. Press Enter to continue."

raw_input ()

data=pandas.read_csv(r'billing_list.txt',sep='|', skiprows=[0,1,2,3,4,6], skipfooter=3,engine='python' )
data['Account ']=data['Account '].astype('str')
data['DocumentNo']=data['DocumentNo'].astype('str')
idx=0

for row in data.iterrows():
    cust_number=data.values[idx][3]
    str_billdoc=data.values[idx][5]
    print cust_number+'-'+str_billdoc
    f1 = open ('vf03_'+ str_billdoc +'.vbs','w')
    f1.write(str_head)
    f1.write('session.findById("wnd[0]/usr/ctxtVBRK-VBELN").text = "'+str_billdoc+'"' + '\n')
    f1.write('session.findById("wnd[0]/mbar/menu[0]/menu[11]").select' + '\n')
    f1.write('session.findById("wnd[1]/tbar[0]/btn[37]").press' + '\n')
    f1.write('session.findById("wnd[0]/mbar/menu[2]/menu[1]").select' + '\n')
    f1.write('session.findById("wnd[0]/mbar/menu[3]/menu[5]/menu[2]/menu[1]").select' + '\n')
    f1.write('session.findById("wnd[1]/tbar[0]/btn[0]").press' + '\n')
    f1.write('session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\work"' + '\n')
    f1.write('session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "'+cust_number+'-'+str_billdoc+'.txt"' + '\n')
    f1.write(str_tail)
    f1.close()
    vbs_name= 'vf03_'+ str_billdoc +'.vbs'
    os.system(vbs_name)
    
    idx=idx+1
 
os.system('del *.vbs')
print "All Billing Documents have been exported."
print "Press Enter to Exit program."
raw_input()
 