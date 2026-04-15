This repository contains the code used for the Magdalene Houses' impact report generator and is linked to the website (not listed for privacy reasons).

**App.py:** Runs the website and generator by taking in the requested year and workbook and producing plots and calculating statistics based on that input. It returns the generated PDF as an output for staff.

**Requirements.txt:** Lists the required packages and tools for Render to work. If you are to import any package that is not already in the app.py script, you **MUST** add it to this text file for it to work.

**Index.html:** HTML for the website itself.

**Report_template.html:** HTML for the report. Uses plot and statistic variables created in the app.py script along with the JPGs in the Backgrounds folder.

**Error.html:** Simple HTML file in the event of an error occurring. Provides hints to what might be the cause of the problem.

**Report.css:** Handles the positioning of each variable, along with the fonts and colors of the captions used in the Report_template.html file.


*Important Notes:*
- If you decide to alter the dataframe when making changes to the report, be very careful, as the plots are created using specific column indices. Make sure you are using the correct column number if you decide to replace a plot or statistic. **Keep this in mind when adding/deleting questions to the Dashboard Spreadsheet as well!**
- Please, please do not create a new Render website with this code! Request collaborator access if you need to make edits, and contact me if you experience any deployment issues.
