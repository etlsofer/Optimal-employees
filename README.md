# Optimal employees
Embedding software works optimally. Including - taking data from the email, automatically sending an email with the results and alerts on work problems

#To run the software you need to run the entery.py

#Table of Contents

1. Introduction
2. A general explanation of the software
2.1 Part I
2.2 Part II
2.3 Gadgets
3. Future goals

#introduction:

To understand the operation of the software, a general understanding of the role of the physical distribution coordinator is needed and thus describes how the software integrates and allows orderly and convenient work for the coordinator.

Under the coordinator there are 2 main roles: the first of which is sending and receiving print orders from the association and its affiliates (office managers, coordinators, and faculties) and the second is managing a distribution system of about 11 employees whose job is to hang posters and occasionally work in extra inventory management jobs, Hanging party supplies, etc.
The first part is very technical and requires mainly a neat record of incoming and outgoing orders, paying factors, type, and cost of work. It is also necessary to know the status of the job at any given moment to enable effective follow-up (for example - whether the work has reached the customer or is it in print).

The second part requires not only management ability but also memory and status for posters (such as hanging date, removal, and whether the distribution was performed), efficient and fair shift inlay for everyone, and constant monitoring of the status of the boards.
We will expand a bit on shift pick-up - at the Technion, there are currently 6 active distribution areas and 2 more areas that are correct for writing the JNF. These are - upper dormitories, lower dormitories, upper and lower faculties, village courses, and stations.
To ensure proper inlay, the coordinator must take into account two main parameters that are the place the employee knows (each employee must undergo an overlap before working) and the time the employee can work. Given these two parameters, there is also room to exercise judgment and prioritize an employee who has filed more shifts but these are not part of the constraints.
Now let's move to the software

#Part a

When the software starts, the main screen will appear
When the Enter hours button is pressed you can enter the employee's salary by the employee's name, the number of hours he worked, the payer load, the reason for payment (distribution round, warehouse arrangement, etc.). There is also an optional field to which you can add additional notes. Automatically enter the feed date so that the time worked by the employee can be restored.
The above data is automatically recorded into the XL file and the software remembers the data.

The add order button will display the following-
You can enter a new order that needs to be made by order name, order date, supplier name, order size (e.g. 3A), arrival date, the number of items, paying factor and order price (currently optional and can be updated later in case of urgent orders) You can also add comments about the order.
The above data is also automatically saved into a file and the software remembers its existence.



The Start new month button starts a new month and saves the previous month's bookings and wages. In addition to archiving the data, it automatically sends the files to the office administrator.

The coordinator needs to confirm that he/she is interested in a new month to prevent a click error
 
Confirmation of sending the email
 

Illustration of the email that comes to the header and file manager (the software knows which month the arrangement refers to) -
 
The above buttons contain most of the technical part of the first part of receiving and sending orders, but there are other options that I will briefly present:
Removing an order from existing orders, changing the status of an existing order, and viewing existing orders in the system (of the current month).
It is also possible to delete the employee's hours and view the hours of existing employees in the system.



The above options are at the top bar





Future goals to add - Sending an email to a vendor through the software from an existing vendor list and choosing the printable file so that to send a work bill, you must enter the work into the system and no forgetting condition can occur.

#part b

Now we will move to the second part which is responsible for managing the employees
Work in the software's contact bar under the tab Work to invite a new employee to the system by clicking on Add a worker Opens the  window -

Adding an employee is entering his name, phone number, email address, and selecting an area to overlap with. Now the employee is added to the system for any future need.
It is also possible to delete an existing employee from the list of employees in the system or add a new area to which he is working, or add the page for an employee who does not want to work in a specific area.
Note - When an employee enters hours, the employee may never be entered afterward because the system does not know this because there are additional jobs that are not related to the ad's suspension of the times the coordinator needs to enter.
In the Send Availability tab, you can choose month and day of the week in the workplace (the new way of working is every Wednesday and the next month for distribution is April)
 
After selecting the above, all employees will send the XL file.
 

In the file, the employees mark their names on the date they can and return the file by email (history is attached by email).
You can download the employee file and the software will download all available files into this dedicated folder and let you know how many employees sent an availability file (the software policy is that when an employee does not send an availability file it is charged on all days)
 
You can delete old files by deleting an old employee file so that no new arrangement can be made for the current month.
Now go to the fourth button in the software that appears on the main screen Upload name (also possible under File / Open) Clicking on it will lead to the next screen -
 
Upload from the file. We can select all downloaded files and extract all employee days data. Load from the keyboard we can enter employee data manually (in the general case there is no work arrangement file but less useful for the needs of the coordinator). Now we can mark whether we want to print the result or an XL file.
The software calculates the optimal arrangement for the employees without the following parameters -
An employee who submits more gets more, each employee gets one day off, who feels the preferred time for a late submission.
The result will be printed to an XL file that looks like this -
 
The above file can be sent back to the employees by clicking Send Monthly Arrangement Well complete the work arrangement, also the employee will receive an SMS warning that the arrangement is ready.
Note - The system also has an automatic result mechanism that makes file loading and printing to the file automatically and fast.

#Gadgets

The software has an alert mechanism that can alert you to problems such as the disadvantage of working groups or overlap with the area.
 

Future Goals - Identify an employee who has not finished his job on time and be alerted.

#Future Goals

- Built-in price list in software that knows how to price each item by quantity, size, and type of order.
- Adding posters that are intended for distribution (currently files exist but need to be integrated into the software) and their work status.
- Submit a weekly work arrangement containing the information about the posters they need to hang and for those on the boards.
- View the current state of the boards - Every employee snaps the board after hanging, all it has to do is connect the software to the data so the boards can be quickly audited (added in the enclosed article that algorithms could be used to identify the location of the board that was shared and added).
