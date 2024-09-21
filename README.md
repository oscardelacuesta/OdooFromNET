# OdooFromNET
Development of a .NET desktop application for partner management in Odoo using XML-RPC. (PoC)
The term CRUD represents the four basic operations that can be performed on data in a management system: Create , Read , Update , Delete . These are the fundamental operations in any system that manages data, and are essential for maintaining and manipulating information in any database or storage system.

In this project, we implement a CRUD system for partner management in Odoo using a VB.NET application , where the following operations can be performed:

Create new partners.
Read (load) the list of existing partners.
Update the information of a selected partner.
Remove an existing partner from the database.

Technologies used
This project combines several technologies and techniques that enable interaction between VB.NET desktop application and Odoo system:

VB.NET : Programming language used to develop the desktop application that will manage the partners, the project is in Winforms.
XML-RPC : Protocol that allows remote communication between systems. Through XML-RPC, the VB.NET application communicates with Odoo to perform CRUD operations in real time. XML-RPC sends requests and receives responses in XML format.
Odoo : The ERP system where partner data is stored and managed.
EPPlus : .NET library used to export data to Excel files (.xlsx).
Windows Forms : Graphical platform for creating user interfaces in Windows desktop applications.

Project Flow
Authentication in Odoo : When the application is started, authentication is performed with the Odoo server using XML-RPC. If the authentication is successful, the system returns a UID that allows the following operations to be performed.
CRUD in the partner list : The user can:
Create a new partner: Enter the data (name and email) and save the record in Odoo.
Read Partner List: The partner list is loaded and displayed in a DataGridView .
Update a partner: Select a partner from the list, modify the data and save it.
Delete a partner: Select a partner from the list and delete it from the system.
Export to Excel : The option to export the list of partners to an Excel file is provided. The user can select the location to save the file and the application will automatically open it after the export is complete.
Conclusion
This project is an excellent example of how to connect a VB.NET desktop application with an Odoo server using XML-RPC . Through this implementation, we were able to perform CRUD operations on partner data and handle errors appropriately to improve the user experience. Additionally, the Excel export functionality allows the user to easily maintain a copy of the data.

Remember that I have used standard Windows controls. Using controls such as Syncfusion or Devexpress among others, you can do wonders with the data obtained, from reports to highly advanced grids, as well as gateways for communication and automation of applications with other ERPs .


www.palentino.es




