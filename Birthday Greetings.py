import datetime
import pandas as pd
import win32com.client as win32

# Load the employee data from an Excel file
df = pd.read_excel('YourExcelLocation.xlsx') # here i set the location of the file that ill be reading

olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

def email_func(subject, birthday_receiver, name):
    mailItem = olApp.CreateItem(0)
    mailItem.BodyFormat = 1
    mailItem.To = 'marcorob@email.com' # here is the email of the recepients.
    mailItem.Cc = email
    mailItem.Subject = subject + ' ' + str(name) + '!' + ' ' + str(month) + '/' + str(day)
    mailItem.htmlBody = (
        "<h1>Today is " + str(name) + "'s Day</h1>" +
        "<img src='Locationoftheimage.png' alt='birthday'>" + # here is the location of the image that im sending
        "<img src='Mysignature.png' alt='signature' width='300' " + # here is my signature or company logo
        "height='200'>" +
        "<style>" +
        "h1 {" +
        "text-shadow: 1px 1px;" +
        "text-align: center;" +
        "font-family: sans-serif;" +
        "font-size: 40px;" +
        "color: navy;" +
        "}" +
        "</style>"
    )

    mailItem.Send()

today = datetime.date.today() # Today day
year = today.year

# and in this loop i go trough each employee to see who match today birthdays.

for i in range(0, len(df)):
    month = df['Birth_month'][i]
    day = df['Birth_day'][i]
    name = df['Name'][i]
    email = df['Email'][i]
    birthdate = datetime.date(year, month, day)

    if birthdate == today:
        email_func('Happy Birthday', email, name)
        print('Send email')
    else:
        print('We didn't send anything today')
