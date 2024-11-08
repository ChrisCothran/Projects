This workbook is intended to provide a means to view the contents of an entire Worksheet object in Microsoft Excel, eliminating the need to scroll horizontally and vertically to view and modify data. 
It consists of two tabs by default, the first (labeled Capacity Data for the default setup) contains the data that the viewer will read and/or modify.
Click the Button on the Viewer tab to activate the Interactive Userform. By selecting an entry from the ComboBox at the top, the corresponding Unique Id loads the associated row in Excel.
This information can be modified and then saved via the column.
An associated ID is read from column 14. If this ID matches for multiple entries, then those facilities are listed as 'Associated' facilities at the bottom of the form. In the LNG context, for example, this could allow five separate production trains to be listed, and grouped by that unique ID.
To modify the Userform, click the Visual Basic Editor button in the Developer Ribbon, then click UserForm1 under the Forms folder in the VBA Project pane.
Any modifications to the name or adding/deleting fields can be entered into the Private Sub CommandButton_click() section.


