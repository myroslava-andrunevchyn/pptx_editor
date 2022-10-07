# pptx_editor
**pptx_editor** replaces text and inserts images in pptx template file with data retrieved from MySQL database. Let’s assume you would like to customize report for each your client. You have database with clients data (something like the name of client, their contact info, logo, etc.) and need to adapt pptx file for each client with certain information inserted into the file.

I’ve used **python-pptx** library (https://github.com/scanny/python-pptx), which is a Python library for creating and updating PowerPoint (.pptx) files. My project consists of the following steps:

1) create and populate MySQL database

2) create pptx template with tags, which are supposed to be replaced

3) write a script which fills in template with data

![Workflow](https://user-images.githubusercontent.com/54477002/194352329-b94631a2-cd10-4c28-b37b-86b98c50011a.png)

The general workflow is following: 

1) we retrieve the list of clients from database(assuming we do not have a list beforehand);
2) having client *id* we retrieve all necessary data from database and make 2 dictionaries of values, where ***key : value*** pair corresponds to ***tag*** in our template and ***information to put in placeholder***

Here is the example of my dictionaries:

dictionary of string values: {'{{client_name}}': 'CompanyA', '{{budget}}': 100, '{{contact}}': 101010, '{{date_est}}': datetime.date(2011, 11, 11), '{{mail}}': 'companya@mail.com'}

dictionary of graphic values: {'{{logo}}': 'C:\Local Storage\Logos\CompanyA.png', '{{city_picture}}': 'C:\Local Storage\Cities\CompanyA_Paris.jpg', '{{diagram}}': 'C:\Local Storage\Diagrams\CompanyA_yellow.jpg'}

1) we iterate through placeholders, search for tags and fill them in with corresponding data

2) save the file with company name
