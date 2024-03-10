# Goal:
# Send Emails via outlook appl;ication using python 
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail= outlook.createItem(0)
mail.To = 'Receiver Mail Address'
mail.Subject = 'SUBJECT OF THE MAIL'
mail.HTMLBody = 'Create a body of the mail using HTML syntax'
mail.Display() # to display the Created mail
mail.Send() # to share the mail
mail.exit