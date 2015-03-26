Create an Excel Workbook for importing the D&B Worldbase 1784

1. Create macro enabled workbook

2. Checkout "Z\_GetMostRecentVersions.bas" from Google Code and import as Visual Basic Module

3. Make sure that the "Microsoft WinHTTP Services" library is referenced

4. In the Excel Trust Center check, in the section Macro Settings, the "Trust access to the VBA project object model"

5. Run macro "ReadCodeAndRefTables"

6. Module Z\_GetMostRecentVersions can now be removed from the workbook

7. Correct, if needed, the currency code reference table (see [issue 1](https://code.google.com/p/read-wb-1784/issues/detail?id=1))

8. Save the workbook

9. Close & re-open the workbook to read the 1784

10. Enable macro's and open the relevant fixed width text file