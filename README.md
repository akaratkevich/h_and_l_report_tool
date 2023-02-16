# h_and_l_report_tool
H&amp;L reporting tool in "go"

A quick summary of what the tool does:
-	Gets the ‘report’ date from the user
-	Loads the downloaded copy of the WS (excel file)
-	Calculates back 7 days from the ‘report’ date to create dates to filter the output
-	Filters further on the status Completed/Cancelled
-	Prints out the results to the screen and to a new excel file

Caveats:
- The file path, for the downloaded excel file statically defined
