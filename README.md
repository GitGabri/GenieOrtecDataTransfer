# GenieOrtecDataTransfer
Transferring data from Genie and Ortec reports to excel using VBA


Gamma results to excel V2.0 instructions:
Scope of the program: to copy data from the Genie or GammaVision report files into an excel spreadsheet easily and painlessly.
1.	Make sure macros are enabled when the worksheet is opened.
2.	Make sure the nuclides you are looking for are in the Nuclide_list.csv file, if not add them.
3.	Choose the nuclides you are looking for by hitting the Choose nuclides button and checking the checkboxes on the form that appears
4.	Choose a report file by browsing using the button up top, or simply copy the path to the A2 cell. Alternatively specify the path to the folder which contains all of your report files and the macro will automatically look for them all and copy all the data. (You may also specify the paths by creating a text file (with a list of the paths separated by enter) giving it the extension .plf (path list file).
5.	Then hit Get Data and the data will magically appear underneath.
6.	Each time you run “Get Data” the macro looks for the last cell in the A column and separates the new data from the old one by a row. If you want to erase all previous data just hit the appropriate checkbox on the top of the sheet.

Things to look out for:
Genie:
The structure of the report files the macro looks for: need to print out the header file before the nuclide identification report, and then the MDA report at the end. ISO MDAs and interference reports not supported at the moment. The macro needs text files not pdfs. Looking to make it available for pdfs too, soon.

the reports in Genie don't actually say what reference date was used for the decay correction so you have to know that yourself, and there is no option to not decay correct in Genie (that I know of) so you never know if the activities reported are decay corrected or not. For that reason the program always puts the Genie activity results under Decay corrected activities.

You can change whether the results are placed in the decay corrected column, or not decay corrected column by checking/unchecking the “Genie Data Column Shift” checkbox.

GammaVision:
The GammaVision read routine takes the MDAs from each energy line but takes the activities from the summary table because this is the only place which has uncertainties. The routine checks automatically for decay corrected data in the summary table


Some tips: 
you can create a text file with the paths to the report files separated by "enter" (I think this is a carriage return and line feed in windows?) save it as a text file with the extension .plf (paths list file) and then specify this instead of a report file. Or if a folder contains the report files then you can specify the folder path and the macro will look through the folder for the reports. 
