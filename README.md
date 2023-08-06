View Calendar events

This is python based script using tkinter module for GUI bases application to retrieve the outlook calandar events.
We are going to use Microsoft Graph API to fetch the data from tenant.
(Please check Graph API endpoint and various url's)
Tkinter is useful library when you dont want to get into complexity of GUI based application.
This small application pulls the calendar meetings details froms specific organization supplied in textbox.
You can view the meeting details and also delete any event if want.
This is helpful to view meetings of individuals as well as from resource mailbox accounts.
It provides all require information associated with this meeting, Subject, date/time, Join Link, Organizer name and numberic meeting ID.
This is just a sample script and needs to modify to get specific details.

Please follow below steps to test this application.

Step 1 - create and register application in Azure AD.
Step 2  -Note down the client ID, secret and Tenant ID to input it in code.
Step 3 - open main python script and add client ID, secret and Tenanat ID details.
Step 4 - Your application is ready to go, run the code which will present small GUI to input organizer email, date and time.
Step 5 - click on fetch, once result populated , you can view and delete the event.
