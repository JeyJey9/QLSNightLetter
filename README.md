# QLSNightLetter
Python Script
 
Updates Logging
 
V1
 
- Converting PDF to Excel
- Extract Excel data
- Insert data to specific columns in "master" excels
 
V1.1
 
- 8 AM timestamp check
- Previous Days Change (-1)
 
V1.2
 
- FCPA compatible
 
V1.3
 
- Added support for sparklines - They are now preserved when data is being updated in the excel files
 
V1.3.1
 
- Information popup for "yesterday value"
 
V1.3.2
 
- Added "PDF Insufficient Data Report"
 
V1.5
 
- Added "Data Shifting"
 
V1.5.1
 
- Reordered actions. Now if you choose to shift data and also to update values the script will first shift data from right to left
   in the set of data you specify and leave the last column empty after which it will paste data inside that empty column it left.
 
V1.5.2
 
- Added a logic for the script to insert the values from the PDFs to every sticker it finds in the sheets.
   + What used to happen is that if a sticker was hidden, it was an issue that was relevant then it wasn't necesarry to have it in so it was hidden,
   + and added again later when the issue became relevant again the script would only update the first occurence of the sticker in the list.
   + Now it updates all occurences on a specific sticker no matter if there are any hidden ones or not.
 
V1.6
 
- Added "broken PDFs" functionality.
   + If a PDF file was generated wrong the relevant data should still be grabbed from it.
 
V1.6.1
- Added "broken PDFs" debugging.
   + In case some PDFs STILL fail, the list that fail will be shown at the end of the script so they can be manually updated.

V1.6.2
- Implemented "Use Yesterday's Daily" logic inside the "Last daily" and "prev daily" checkboxes.

V1.6.3
- Added warning if "Study_NL" excel is in "read-only" mode.
   - Sometimes an excel process might still be running in the background from failing to exit the 
   - master excel file which results in data not being updated at all.