
Set obj = createObject("Excel.Application")   		'Creating an Excel Object
obj.visible = false                                	'Making an Excel Object invisible
'Set obj1 = obj.Workbooks.open("C:\Users\HassanH\Desktop\IERS\IERS Software\IERS.xlsm") 

'Set obj1 = obj.Workbooks.open("IERS.xlsm") 

Set WshShell = CreateObject("WScript.Shell")
strCurDir = WshShell.CurrentDirectory
Set obj1 = obj.Workbooks.open(strCurDir & "\IERS.xlsm") 

'obj1.Close                                              
'obj.visible = true					
obj.Quit                                                'Exit from Excel Application
Set obj1=Nothing                                        'Releasing Workbook object
Set obj=Nothing                                         'Releasing Excel object

'Set obj2=obj1.sheets.Add                               'Adding a new sheet in the excel file
'obj2.name="Sheet1"                                     'Assigning a name to the sheet created above
'Set obj3= obj1.Sheets("Sheet1")                        'Accessing Sheet1

'obj3.Delete       'Deleting a sheet from an excel file
