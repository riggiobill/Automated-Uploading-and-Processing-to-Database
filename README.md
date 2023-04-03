

# Automated-Uploading-and-Processing-to-DatabaseExcel.
Set up an easy automation command for cleansing, uploading, and processing daily data for an international shipping company client. 

![alt text](https://github.com/riggiobill/Automated-Uploading-and-Processing-to-Database/blob/main/Screenshots/Excel.png?raw=true)


![alt text](https://github.com/riggiobill/Automated-Uploading-and-Processing-to-Database/blob/main/Screenshots/UploadPy1.png?raw=true)




To perform this task, I built a Python script that cleanses a daily update of data, connects remotely to a digital database service, and then uploads and processes the data. This was automated through the use of a .bat file as a service.


## Step 1 - Cleansing and Preparing Data

* Access and edit the raw, daily update spreadsheets of data using Openpyxl to remove unnecessary sheets and confirm the names of the remaining sheets for uploading.

![alt text](https://github.com/riggiobill/Automated-Uploading-and-Processing-to-Database/blob/main/Screenshots/UploadPy_confirm-save.png?raw=true)


## Step 2 - Define send_email function to use for error reporting

* Defines a "send_email" function using smtplib to create a function which expects basic email and login info.
* Declares a set of user info variables which can be altered as per client needs and email recipients. Passes in the password from a separate imported file.
* Will exit automation on detection of an error, and will return an error message notification email.



## Step 3 - Confirm all renaming, saving, and file locations of cleansed spreadsheets.

* Checks the titles of both files and the worksheets in the files, which will need to match the expectations for the uploading and processing on the database side.
* Checks to confirm the removal of unnecessary extra worksheets in the raw data.
* Renames the titles of the spreadsheets and saves them with their new expected titles using os.path.isfile, and double-checks to confirm the titles in preparation of uploading.




## Step 4 - Check all columns and data types before Birst upload.

* Uses Column Name as a value for comparison against expected column names, returns an exit and an error if a discrepancy is detected - performs the same check for cell.data_type as well.
* Checks for each expected column and data type pairing, for one spreadsheet and then for the other. Since the code exits on error detection, if it makes it to the end of the checks the data is cleansed and confirmed to be ready for upload.

![alt text](https://github.com/riggiobill/Automated-Uploading-and-Processing-to-Database/blob/main/Screenshots/UploadPy_data-checks.png?raw=true)

* Certain exceptions are provided for in the code - namely for data that is currently blank, and will not be available yet but will be added to the sheets in the future.
 

## Step 5 - Define a workflow which can connect to Birst and call a webservice command.

* Space details are listed, pulling the password from an imported separate file. These can be updated here as required, as used as stable variables throughout the rest of the code.
* Defines a class "workflowClass" to create and run the workflow, and to associate variables with it.
* Defines functions for space initialization, creating a client, logging in with the login info provided, logging out, and executing a workflow. Also included is a poll workflow function to test and interact with the status of the workflow, determining retries and errors.

![alt text](https://github.com/riggiobill/Automated-Uploading-and-Processing-to-Database/blob/main/Screenshots/UploadPy_workflow1.png?raw=true)

* run_workflow ultimately attempst to run the named workflow, retrying if possible, pausing between runs, and reporting an error on failure.


## Step 6 - Create an instance of the workflow class and use it to run all necessary workflow functions outlined above.
* Creates "x" as an instance of a workflow as defined above. Proceeds to create client, login, run_workflow, and logout.

![alt text](https://github.com/riggiobill/Automated-Uploading-and-Processing-to-Database/blob/main/Screenshots/UploadPy_workflow2.png?raw=true)

* If it reaches this part of the code, it prints a message announcing that the workflow was run, announcing the time of the run command, and finally announcing that the Upload.py file has finished its process.

