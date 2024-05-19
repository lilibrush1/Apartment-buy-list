# Apartment-buy-list
Google sheets code to distribute roommate purchases for shared spaces.

To use this, create a new Google Sheets document. Then, go to Extensions -> App Script. Make a new file and paste the code from code.js into it. Now, go to the clock icon on the lefthand menu that says "Triggers". Create a new one with the "+ Add Trigger" button. Keep everything at default except the following: 
  - Select "copySheetDataWithFormatting" as the function to run.
  - Select "On change" as the event type
  - Change your failure notification setting if you don't want to be notified.

How to use:
Create a sheet called "Summary". Then, create sheets for each category you want and for each roommate. Then, go into the code and make some edits:
  (1) Change the variable on line 3 (sourceSheetNames) to match the names of the category sheets you created.
  (2) Change the variable on line 15 (nameKeys) to include all roommate names paired with their desired color.
  (3) Change the variable on line 48 (nameColors) to include all roommate names paired with their desired color.
Now press the play button that says "Run" to its right.

What this does:
Now, you can edit the sheets for each category. For example, in mine I added items to the "Kitchen" sheet that we needed to buy. I saw that I already had a pot, so I changed that cell to my color. Now, when I go to the "Summary" sheet, everything I added is there and the pot's cell is highlighted in my color. Further, the sheet with my name has the pot written on it too so I know to buy it.
