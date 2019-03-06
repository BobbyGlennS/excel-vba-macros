# Excel VBA macros
These Excel VBA macros are written to do some pre-processing of data on Excel files. They are written specifically for research data of the MySweetheart cohort study. However, the functionality is quite general purpose so these macros can be used for other datasets as well.

## Variable Selection Macro (*variable selection macro.bas*)
### What does it do
This macro is aimed at Excel files that contain many columns, of which only a few are of interest.

The macro creates a new worksheet in the Excel file containing only those variables which are specified by the user, and in the order that they are specified. You can specify variables through their names, assuming that the first row of each column contains its name.

### How to use it
This is a step-by-step guide to get the selectColumnsByNames() macro working.

1. Open the relevant Excel file
2. **Save** it **as** an Excel macro enabled workbook (\*.xlsm)
3. Create a new sheet, named *filters* (make sure the name matches exactly, the macro is case sensitive).
4. The new *filters* worksheet is where the macro will look for the names of the variables you want to keep. It will look for variables in column A. To this end, give column A an appropriate name in cell A1, e.g. *variables_to_keep*.
5. Then add, in row A2, then A3, etc., the names of those variables you want to keep. Make sure the names match exactly those in the data, in particular be aware of spacing and casing.
6. Right click on one of the worksheets and choose **View Code**
7. Choose **insert** -> **module**
8. Copy paste the contents of the *variable selection macro.bas* file into the module window.
9. Now you need to make two small edits in the macro itself: First, find the line that says *dataSheet = "worksheet_with_the_data"* and replace *worksheet_with_the_data* with the name of the worksheet that contains your data. Again be careful that the name matches exactly, and make sure the quotation marks are still there.
10. Then, find the line underneath saying: *newDataSheet = "selected_variables_sheet"*. This defines the name of the worksheet where the subset of selected variables will be stored. If you wish for this worksheet to have a different name (for example in the unlikely event you already have a sheet named like this), then you can rename it here.
11. Now you're ready to run the module. To do so, press **F5**.

If everything has gone well, a new sheet will have appeared in your Excel file (as named under step 10), containing the variables you asked for.

Note that the macro will execute regardless of whether variables you specified are correct or not. If a variable you specified does not exist a column with the variable name will still be added to the new sheet, however the variable name will have "DOES_NOT_EXIST" appended to it. There will be no data in this column. If a variable is not present it is most likely not there in the original data, either not present at all, or there is a typo in the list of variables that you specified.

Top tip: Don't ever run the macro (or any macro) on your original data, always first save a copy. The actions cannot be undone, and as such you may lose data forever if you're not careful.
