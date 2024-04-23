# Microsoft Office AIP Labelling Activities for UiPath
I have developed custom activities for getting and setting sensitivity label on MSO application files. For now, this activity is compatible with Excel (*.xls*), Word (*.doc*) and Powerpoint (*.ppt*) files. Outlook mail labelling will be implemented as a future enhancement.
 
Sensitivity label can be set across different file types (i.e. a label you retrieve from an Excel workbook can be used to set on a Word document or Powerpoint presentation.)
 
### There are 2 activities in this package. 
GetSensitivityLabel - This activity retrieves sensitivity label information from an already pre-labelled file.
SetSensitivityLabel - This activity sets a sensitivity label on a target file.
### Compatibility
UiPath Studio version >= 2021.10.6 (Windows Project, .Net6)
UiPath Studio version < 2021.10.6 (Windows Project, .Net5)
UiPath Studio (Windows Legacy Project, .Net Framework 4.6.1)
 
## Recommendation for usage
It is not necessary to use the "Get" and "Set" activity together. 
 Since all label arguments are strings, you can get the label information during development and store it in a mapping file or any other data structure. If there are any changes to the label information, you can simply re-fetch and update the mapping file thereafter. 
If you are expecting label information to change from time to time, you could use the "Get" activity during Initialization and save it into the Config dictionary or any other variable instead of retrieving it for every transaction. 
