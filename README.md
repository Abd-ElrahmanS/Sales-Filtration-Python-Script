Explanation and Documentation of The Python Script
How the Script Works:

1.	File Handling and Organization:
•	The script begins by scanning a designated source directory for Excel files that need processing. Each file is identified by specific naming conventions ('Report.xlsx' and 'Final.xlsx').
•	After processing, the files are moved from the source directory to a destination directory to maintain organization and avoid reprocessing the same files.
2.	Data Processing:
•	For each file, the script reads the data into a panda DataFrame, allowing for powerful data manipulation. This includes filtering the data based on certain management levels and other specific criteria provided in a separate Excel file, which contains the names and email addresses of the report recipients.
3.	Report Customization and Generation:
•	Once filtered, the data is further customized, potentially involving formatting changes such as converting numerical values to percentage formats.
•	The filtered and formatted reports are then saved as new Excel files in a filtered directory, named uniquely to prevent overwriting and to facilitate easy identification.
4.	Automated Email Distribution:
•	For each newly created report file, the script sends an email to the appropriate recipient as listed in the external Excel file. The email includes the report as an attachment, ensuring that specific stakeholders receive the data relevant to their operational needs.
5.	Logging and Error Handling:
•	Throughout its operation, the script logs its activities, providing clear feedback on the processing of each file and any errors encountered. This makes troubleshooting and verification easier for users.
