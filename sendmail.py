import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'someone@somewhere.com'
mail.Subject = 'Message subject'
mail.Body = 'Message body'
#mail.HTMLBody = '<h2>HTML Message body</h2>'# this field is optional

#In case you want to attach a file to the email
#attachment  = "Path to the attachment"
#mail.Attachments.Add(attachment)
mail.Send()
