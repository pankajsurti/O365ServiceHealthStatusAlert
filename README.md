Office 365 Service Health is a reporting and alerting solution. Due to its design and functionality, I like to refer it as Solution. It uses Office 365 Service Communications API.
It uses PowerShell, HTML and CSS for conditional formatting. We can generate a HTML output report or schedule it to run on periodic intervals with the help of Windows Task Scheduler. 

We can also add small snippet of code at the end to Send Email Alert Notifications. Though Office 365 Service Health does not give tenant specific information. This solution can be modified later to incorporate new functionalities that will be rolled out by Microsoft - Like to include user count of an affected tenant.
The Solution uses JSON Config file to load the configuration like Application ID, Client Secret, AAD Instance, TenantDomain and Log file path. 

To obtain Application ID and Client Secret we must first register an App in Azure AD and then enable Service Communications API. 
