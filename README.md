Steps to run locally:
1. Create project in google cloud console
2. enable google slides and sheets api
3. create oauth2.0 client id in google console and download the client_secret.json file and replace the current one in the repo with it
4. replace line 28 with the new spreadsheet id from the google form responses(https://docs.google.com/spreadsheets/d/1ze6RSxdDzTyK3Zo7wN9NPm6rDCnjft76d3SzUaJq3vw/edit?pli=1&gid=1399428418#gid=1399428418, the id is 1ze6RSxdDzTyK3Zo7wN9NPm6rDCnjft76d3SzUaJq3vw for example)
5. replace the "ranges" parameter in lines 29 and 35 with the name of the sheet within the spreadsheet(found in bottom of spreadsheet)
6. create slideshow manually in google drive, this will be added to programmatically when the code is run
7. Update desired id of slideshow to be generated in line 39
8. Update the values and indices in thw while loop starting at line 61 to match the values in the spreadsheet
9. run code with python slideGen.py command


Things to note:
1. if the picture uploaded from the spreadsheet is a .HEIC file(live photo), the code will crash. Google slides api does not support live photos. Take note of the people who have these(the last row to be processed will be printed when the code runs) and create their slides manually(a pain but no workaround)
2. The code runs a max of ~50 rows at a time. so you will have to update the value in line 58 for where you want(i should = the row number in the spreadsheet that you want to start at, - 1
