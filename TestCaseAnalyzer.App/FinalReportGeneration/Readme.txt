Step.1
Install .NET Runtime 6 (Only first time install)
https://dotnet.microsoft.com/en-us/download/dotnet/6.0

Step.2 
Upload HTML report into folder 'Input'
HV -> L1
Fusa -> L2
Full -> L3

Step.3 
Press 'FinalReportGen' (.bat)

Step.4
Final report will be in 'Output' folder


-------------------------------------------------------------------------------------------
-------------------------------------------------------------------------------------------

# Final report generation
After report generation is done, the result will be shown in 4 possible types
* PASSED 
* FAILED
* OBSOLETE
* JUSTIFIED
* No HTML and Reviewed Needed

## Logic behine the output
PASSED -> Result from HTML
PASSED with comment -> It means HTML test result is passed but there is number of "Not execution" & "Failed" 
FAILED -> Result from HTML
OBSOLETE -> No HTML & Test spec result is OBSOLETE
JUSTIFIED -> HTML 'PASSED" but Test spec 'JUSTIFIED' with comment
No HTML and Reviewed Needed -> File name and HTML ID are mismatch -> This case tool will take file name as ID -> If this ID does not belong in Test spec, this output will be shown.

## How to review the report
User need to validate by changing 'Test Result' , 'Test Result Details' and Add KAP ticket number

These are the following cases.
### No HTML and Reviewed Needed

### NOT TESTABLE
If it's 'NOT TESTABLE' -> fill 'The number of "NotExecuted'

### FAILED
Check history if it was 'JUSTIFIED' -> change column 'Test result' to 'JUSTIFIED' and move 'The number of "FAILED"' to 'The number of "JUSTIFIED"'
If it's failed with ticket, user need to add ticket number

### PASSED with Verification Coverage Comment
It means HTML test result is passed but there is number of "Not execution" & "Failed" 
User need to check the history whether to change test result from "PASSED" to "JUSTIFIED"

### JUSTIFIED
User need to verified whether this TC need to chang to be 'PASSED'