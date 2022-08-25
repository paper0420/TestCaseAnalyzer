#### Step.1
Install .NET Runtime 6 (Only first time install)
https://dotnet.microsoft.com/en-us/download/dotnet/6.0

#### Step.2 
Upload HTML report into folder `Input`
* HV -> `HTMLReportL1`
* Fusa -> `HTMLReportL2`
* Full -> `HTMLReportL3`

#### Step.3 
Run `FinalReportGen` (.bat)

#### Step.4
Cmd will pop up
* Enter: Carline (`G60`,`G70` etc.)
* Enter: Report type (`HV`,`Fusa`,`Full`)
* Enter: SW Release
And Enter

#### Step.4
Final report will be in `Output` folder


# Final Report Generation
After report generation is finished, the result will be shown in 6 possible types
* PASSED 
* FAILED
* OBSOLETE
* JUSTIFIED
* NOT TESTABLE
* No HTML and Reviewed Needed

## Logic behine the output
- PASSED -> Result from HTML
- PASSED with comment -> It means HTML test result is passed but there is number of "Not execution" & "Failed" 
- FAILED -> Result from HTML
- OBSOLETE -> No HTML & Test spec result is OBSOLETE
- NOT TESTABLE -> No HTML & Test spec result is NOT TESTABLE
- JUSTIFIED -> HTML 'PASSED" but Test spec 'JUSTIFIED' with comment
- No HTML and Reviewed Needed -> 
    - Case 1: File name and HTML ID are mismatch -> This case tool will take file name as ID -> If this ID does not belong in Test spec, this output will be shown.
    - Case 2: No HTML but this is  valid test case ID (There is test result & valid requirement in Test spec file)

## How to review the report
User need to validate by changing `Test Result` , `Test Result Details` in order to correct total number of tables `Catagory` and `Detailed cases` 

These are the following cases.
1) No HTML and Reviewed Needed
- Case 1: Check file name & HTML header (Fix from Canoe)
- Case 2: Check why this test case was executed.
    - if this test case is justifie, change `Test Result` and fill `The number of "NotExecuted"`.
    - if this test case is unused, it should be removed from Test spec file.

2) NOT TESTABLE
If it's 'NOT TESTABLE' -> fill 'The number of "NotExecuted'.

3) FAILED
- Check history if it was 'JUSTIFIED' -> change column `Test Result` to 'JUSTIFIED' and move `The number of "FAILED"` to `The number of "JUSTIFIED"`.
- If it's failed and ticket need to be opened. -> User need to put KAP ticket in column `Problem Management Ticket`.
    - **FuSi_Issues need to fill manualy**.
    - **Functional_Isseus need to fill nanualy**.


4) PASSED with Verification Coverage Comment
It means HTML test result is passed but there is number of "Not execution" & "Failed". 
User need to check the history whether to change test result from "PASSED" to "JUSTIFIED".

5) JUSTIFIED
User need to verified whether this test case need to chang to be 'PASSED'.