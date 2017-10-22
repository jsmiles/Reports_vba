# Reports_vba
A very hacky way to pull some data from an Office for National Statistics excel file.

# Contents
- emp01sasep2017: the base data used for this example
- emp01sasep2017_results: what the end results looks like (i.e. it includes 1 extra tab with some data extracted)
- ons_report_vba: the workbook that contains the vba code
- ONS_Output_module: the vba module exported

# How do I use this?
1. Download the "emp01sasep2017" and "ons_report_vba" files and store them locally
2. Open both workbooks in Excel
3. Within the emp01 file, Ensure that you "enable content" when you open it.
4. Within emp01, check the developer tab/macro security button and ensure you enable all macros
5. Within emp01, go to the developer tab and click the macros button
6. There should be a single macro (this macro is visible because you have both workbooks open) visible, click run.

This should have produced a third tab, "My_results" which is populated with four statistics and descriptions. 

# What does it do?
There are four macros used in this process.
1. Creates a new tab in the excel workbook
2. Populates a column with some text that explains the data that we will soon input there
3. Populates the column to the left of the text colum with data that we require. This is the complicated step as all
the data calculations are contained within this step
4. An aggregate function that means
