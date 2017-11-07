# coding:gbk
import pandas
import os
import datetime
import time
import glob
import win32com.client as win32
import re
import codecs
import sys
reload(sys)
sys.setdefaultencoding('gbk')



str_start_date=raw_input('Enter Start Date [YYYY-MM-DD]:')
str_end_date=raw_input('Enter End Date [YYYY-MM-DD]:')
str_date_range = str_start_date[0:4] + '��' + str_start_date[5:7] + '��' + str_start_date[8:10] + '����' +  str_end_date[0:4] + '��' + str_end_date[5:7] + '��' + str_end_date[8:10] + '��'
#print str_date_range
list_bill=glob.glob("C:\\work\\8*.txt")
 

cust_list=[]

for bill in list_bill:
    stra= re.findall(r'80.*[txt|TXT]',bill)[0]
    #print stra
    if stra[0:8] not in cust_list:    cust_list.append (stra[0:8])
    

#print    cust_list 

cust_data=pandas.read_csv('custdata.txt',sep='|',encoding="gbk")
cust_data['SAPID']=cust_data['SAPID'].astype('str')
cust_data['TOADDRESS']=cust_data['TOADDRESS'].astype('str')
cust_data['CCADDRESS']=cust_data['CCADDRESS'].astype('str')
#print cust_data

outlook = win32.Dispatch('outlook.application')



for cust in cust_list:
    #print cust
    
    payment_amount =0.0
    customer_cname=str(cust_data[cust_data['SAPID']==cust].values[0][1])
    to_address = str(cust_data[cust_data['SAPID']==cust].values[0][2])
    cc_address = str(cust_data[cust_data['SAPID']==cust].values[0][3])
    
    print "Processing " + customer_cname
    body_part1=""
    for bill_text in glob.glob("C:\\work\\" + cust +"*.txt"):
        f1 = codecs.open (bill_text,'r','utf-8')
        for i in range (1,25): 
            line = f1.readline()
        temp=f1.readline().strip()
        if temp.find('����') <> -1: 
            sign = -1
        else:
            sign =1 
        body_part1 = body_part1 + '\n' + temp
        
        f1.readline()
        temp=f1.readline().strip()
        body_part1 =  body_part1 +  '\n' + temp
        payment_amount = payment_amount + sign * eval(re.split(r'\s+', temp)[4].replace(',',''))
        
        
    mail = outlook.CreateItem(0)
    mail.To = to_address
    mail.cc = cc_address
    mail.Subject = customer_cname  + '��ʿ����Ʊ�˵�-' + str_date_range 
    mail.Body = customer_cname + '\n' + '��˾' + str_date_range +'�ڼ��˵��總����'
    mail.Body = mail.Body + body_part1 + '\n'
    mail.Body = mail.Body + '\n' + '�����ܽ�' + format(payment_amount, '0,.2f') + 'Ԫ\n'
    mail.body=mail.body + '''
    
�յ���ȷ�ϲ��ظ������߸��������ڲ��ظ�����Ĭ���˵����ݣ������ڷ�Ʊ�˵�������30���ڸ�� 
�ǳ���л������ϣ� 
���� 
�Ϻ�����������԰���޹�˾ 
Ӧ���˿ 
��ϵ�绰��+86-21-2060 4369 
�������䣺SHDR-ACCOUNT.RECEIVABLE@DISNEY.COM 

    '''
    
    #mail.HTMLBody = '<h2>HTML Message body</h2>'# this field is optional
    #In case you want to attach a file to the email
    for attachment  in glob.glob("C:\\work\\" + cust +"*.*"):
        mail.Attachments.Add(attachment)
    mail.Send()
    
print ("All Emails are processed. Please go to Outlook and check.")
print ("Press Enter to close program...")
raw_input()
