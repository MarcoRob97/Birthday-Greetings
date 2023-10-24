# BirthdayGreetings
Automating Birthday Email Notifications for Employees

Have you ever wondered how to automate birthday greetings for employees in your organization? In this post, I'll walk you through a Python script I created to streamline the process of sending birthday congratulations via email.

How it Works:

Importing Necessary Libraries:

We start by importing essential libraries.
datetime allows us to work with dates.
pandas is used to read employee data from an Excel file.
win32com.client enables interaction with Microsoft Outlook.
Loading Employee Data:

We load employee data from an Excel file ('File_Location.xlsx') into a Pandas DataFrame (df). This data contains information such as the employee's name, email address, birth month, and birth day.
Setting Up Outlook:

We create a connection to the Outlook application (olApp) and the MAPI namespace (olNS). These are essential for sending automated emails through Outlook.
Defining the email_func Function:

We define a function called email_func that takes three parameters: subject, birthday_receiver, and name.
This function creates an email message with the provided subject and recipient.
It sets the email's subject line to "Happy Birthday [Name]!" and includes the date (month/day) for added personalization.
The email's body is formatted in HTML, with a festive header, birthday image, and a company signature.
Checking Birthdays:

We get today's date using datetime.date.today() and extract the current year.
We loop through the employee data in the DataFrame to check each employee's birthdate.
For each employee, we retrieve their birth month, birth day, name, and email address.
We create a birthdate variable using the current year and the employee's birth month and day.
The script then checks if the birthdate matches the current date (i.e., it's the employee's birthday today).
Sending Birthday Emails:

If it is the employee's birthday, the email_func function is called to send a birthday email.
The email includes a personalized message with the employee's name, a birthday image, and a company signature.
The email is sent using Outlook.
Logging the Process:

Regardless of whether an email is sent, the script logs the process. If an email is sent, it logs "Send email"; otherwise, it logs "We didn't send anything today."
Customization and Automation:

This script can be customized by modifying the Excel file with your employee data and adjusting the email content.
To automate the process, you can schedule this script to run daily using task scheduling tools or CRON jobs.
By running this script daily, you can ensure that every employee receives a special birthday greeting without any manual effort. It's a practical example of how Python can streamline routine tasks and improve workplace efficiency, saving time and ensuring employees feel appreciated on their special day.
