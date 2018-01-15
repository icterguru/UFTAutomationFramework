
'Main Script
 Dim rc
 Dim i

'Count the total number of executable rows 
 rc = DataTable.GetRowCount
 PRINT RowCount
 'Loop through the total number of executable rows 
 For i = 1 to rc

	'Set i value in the current row
	DataTable.SetCurrentRow(i)
	'
    'Check if Execute is set to Y or y
	If UCase(DataTable.Value("Execute")) ="Y" Then
	
 Reporter.ReportEvent micDone,  "Test Name: " & DataTable.Value("TestName") & "  started...", ""
		'If true then Execute TestName from the corresponding row in the Datasheet
		'The sub routines are located under Library  Functions
 		Execute  DataTable.Value("TestName")

Reporter.ReportEvent micDone, DataTable.Value("TestName") & "  ended.", ""
	End If
Next



