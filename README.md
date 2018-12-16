# GCP_APP_SCRIPT_BILL_SCRAPER
This App Script scrapes Gmails and add the bills to be paid to Excel for Audit 


Currently this app is written to run every day. 
To achieve this a trigger need to be set 


This app currently can manage the following bills

1) City of Toronto Electricity Bill
2) Enbridge Gas Bill
3) Alectra Electricity Bill (Account number has to be hard corderd)
4) Canada Post Epost Bill (Notify when to check canada epost - no value or account number for now)
5) City of Oshawa PUC Electricity Bill 
6) City of Markham Property Tax Bill (Notify when to check canada epost - no value or account number for now)

Following Email configuration file should be available in Google Drive and the link for this file should be updated in get_newmails() method.

Sheet1
Col 1 (Service Provider - Eg: Toronto Hydro) | Col 2 (Email Address - contactus@torontohydro.com) | Col (Subject - 'Your bill is ready')

Sheet 2
Col 1(Property - '123 House') | Col 2 (Service Type 'Electricity' ) | Col 3 (Account Number '123456')

Follwoing audit file which will be updated when a new bill arrive should be created in Google drive and the link should be updated in extract_detail() method. 

(Col1)Date	| (Col 2)Property	| (Col 3)Service Type	| (Col 4) Account Number | (Col 5) Amount	| (Col 6) Status | (Col 7) Notes
