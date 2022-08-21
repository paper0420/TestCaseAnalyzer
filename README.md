# Final report generation
After report generation is done, the result will be shown in 4 possible types
* PASSED
* FAILED
* OBSOLETE
* No HTML and Reviewed Needed

## Review report
if there is not testble or justified test case, user need to validate by changing 'Test Result' and 'Test Result Details'

These are the following cases.
### No HTML and Reviewed Needed
Check history if it was 'NOT TESTABLE' -> change 'Test Result' to 'NOT TESTABLE' and fill 'The number of "NotExecuted'

### FAILED
Check history if it was 'JUSTIFIED' -> change column 'Test result' to 'JUSTIFIED' and move 'The number of "FAILED"' to 'The number of "JUSTIFIED"'

### PASSED with Verification Coverage Comment
It means HTML test result is passed but there is number of "Not execution" & "Failed" 
User need to check the history whether to change test result from "PASSED" to "JUSTIFIED"