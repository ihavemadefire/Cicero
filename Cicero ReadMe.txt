Cicero is a VBA tool used to generate dead meter backbilling cases.



"Cicero External Master.xlsm" - A master spreadsheet that contains all revenue assurences cases.  The list is generated
When a Water meter is replaced.  This flags the account for review.  A table query grabs the Customer Type, Book, Customer 
Reference number, Name, Service address and Property Reference number.  The property reference number is used to generate 
a customer usage report using SAP Crystal Reports.  All of the customer info here is dummy data drawn from only the finest 
TV and Film characters.

Crystal Reports - In the intertests of customer privacy, I haven't included one of these reports.  For our pourposes, make 
up some usage values (1-10 works best).

To work a case, launch Cicero and the External Master.  First, click "Update From Master" in the To Be Worked column.  This 
will pull in the information already in your external master file.  On the dropdown bar, click on "Initial Review" and this 
will poulate the list of all unworked cases.  Click on the first case in the list and much of the review info will be 
populated automatically.  Now, enter in the Previously billed amount, if any (just make it up (1-5 works), The old meter 
Number, the New Meter number, The month that the meter stopped registering, the moth that it was replaced, The meter change
Date, the Mailing address (two sepearate feilds), And then Enter the same months from the previous year and their read values,
i.e. if the customer was out from February 2018 to August of 2018, then B20 - B26 should be Feb-17 - Aug-17.  Then make up 
someusage numbers for those months.

Now Press the Append button.  This appends the values for thaat rate case to the Repository.  In turn, it populates the feilds 
in the review column.  If the numbers look right, Click the "Finish" button. This will append all of the valeus to the table 
and store a copy of the needed files in a folder on the desktop named Cicero.  Notice it also removed that case from the Initial 
Review stack and now is under the out for review stack.  Under normal operation these files are printed out and handeed off 
to managemnt to sign. The print function has been disabled for your benefit. Alternately, if we need to wait for the next read 
to come in, Click the hold button to change the record's status to "On Hold".

Once all the cases have been worked, with the Master file open, Click on the Update Master button.  Save Both. 