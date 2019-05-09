# ICP-scripts
Scripts developed to assist ICP staff

### ExtractAppointments
VBA code for exporting shared calendars from Outlook to Excel. Relies on user's account having read access to the calendars specified within ExtractAppointments

To add the code to Outlook:
1. Start Outlook
2. Press ALT + F11 to open the Visual Basic Editor
3. If not already expanded, expand Microsoft Office Outlook Objects
4. If not already expanded, expand Modules
5. Select an existing module (e.g. Module1) by double-clicking on it or create a new module by right-clicking Modules and selecting Insert → Module and give it an appropriate name (eg. ExtractAppointments)
6. Copy the code from ExtractAppointments and paste it into the right-hand pane of Outlook’s VB Editor window
7. Click the diskette icon on the toolbar to save the changes
8. Close the VB Editor

To run the saved macro from Outlook:
1. Press ALT + F8 to open macro list
2. Select ExportCalendarsToExcel
3. Click run
4. Provide a date range of interest as instructed
5. Click OK
6. Use and/or save the resulting spreadsheet as required
