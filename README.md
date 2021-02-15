# Perkspace
Simple python code to take attendance in google meet\
Here students are asked to enter a secret key (for example '123x'), using selenium we scrape the live chat data from meet and store it as a txt file locally\
If the key is detected that student will be marked present in the excel sheet (Excel sheet with name and register number must be created before hand).\
New columns with current date will be added to the sheet dynamically.\
\
Libraries used:\
- selenium
- os
- re
- pandas
- smtplib
- google\
\
A Sample excel file is added to this repository, make sure the column names are same as this excel file.
