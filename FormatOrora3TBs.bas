Public firstDestinationCell As String
Public destinationFirstColumn As String
Public destinationFirstRow As Long
Public toCell As String
Public sourceDataSheet As String
Public sourceDataRowCount As Long

Sub Main()
	OptimiseVBA True 'Must run this to speed up the process (a bit, this is due to Sum If formula is a large calculation. A N rows TB table means there will be N * N times calculation.)
	
	SetSourceData 'Set some global variables to be used across multiple functions
	
	PopulateFormulas 'customise here
	CleaningUp 'customise here
	
	'MsgBox "All forumlas are populated. Click to continue the calculation" 'Tell user all done
	
	Dim startTime As Date
	Dim currentTime As Date
	Dim duration As Double
	duration = 0
	startTime = Now()
	
	Do While Not Application.CalculationState = xlDone
		currentTime = Now()
		duration = currentTime - startTime
		Application.StatusBar = "Recalculation In Progress: " & Format(duration, "hh:mm:ss")
		DoEvents
	Loop
	Application.StatusBar = False
	
	MsgBox "Calculation Completed. Please save the file."
End Sub

Sub OptimiseVBA(isOn As Boolean)
	Application.Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
	Application.EnableEvents = Not (isOn)
	Application.ScreenUpdating = Not (isOn)
	ActiveSheet.DisplayPageBreaks = Not (isOn)
End Sub

Sub SetSourceData()
	
	destinationFirstColumn = "B" 'Specify which column to fill the value from
	destinationFirstRow = 33 'specify the row number of the column above
	sourceDataSheet = "Full TB" 'Source data sheet name
	
	
	'Do Not Touch below, only change above
	With Worksheets(sourceDataSheet)
		toCell = .Range("A1").SpecialCells(xlCellTypeLastCell).Address 'last cell in source tab
		
	'Get Source Data Last Row
		sourceDataRowCount = .Range("A1", toCell).Rows.Count - 1 'get row count, -1 is to exclude header row
		
		'.Name = "Temp" 'update the name to speed up sumif filling
	End With
	
	'ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
	'ActiveWorkbook.Sheets(Worksheets.Count).Name = sourceDataSheet
End Sub

Sub PopulateFormulas()
	
	FillHeader "Cover"
	
	'do all sheets need data populated
	DoPwcEntityLevel
	DoPwcAccountLevel
	DoClientEntityLevel
	DoClientAccountLevel
	'end of custom actions
End Sub

'this step fills the year/period etc on the selected worksheet
Sub FillHeader(sheetName As String)
	With Worksheets(sheetName)
		.Range("C16").Formula = "='" & sourceDataSheet & "'!$Y$2"
		.Range("C18").Formula = "=VLOOKUP(""*"",OFFSET('" & sourceDataSheet & "'!A2, 0, 0, ROWS(A:A) - ROW(A2) + 1),1,FALSE)" 'Primary TB Name'
		.Range("C19").Formula = "=VLOOKUP(""*"",OFFSET('" & sourceDataSheet & "'!F2, 0, 0, ROWS(F:F) - ROW(F2) + 1),1,FALSE)" 'Primary TB Year'
		.Range("C20").Formula = "=VLOOKUP(""*"",OFFSET('" & sourceDataSheet & "'!G2, 0, 0, ROWS(G:G) - ROW(G2) + 1),1,FALSE)" 'Primary TB Period'
		.Range("C22").Formula = "=VLOOKUP(""*"",OFFSET('" & sourceDataSheet & "'!AH2, 0, 0, ROWS(AH:AH) - ROW(AH2) + 1),1,FALSE)" 'Secondary 1 TB Name'
		.Range("C23").Formula = "=VLOOKUP(""*"",OFFSET('" & sourceDataSheet & "'!AI2, 0, 0, ROWS(AI:AI) - ROW(AI2) + 1),1,FALSE)" 'Secondary 1 TB Year'
		.Range("C24").Formula = "=VLOOKUP(""*"",OFFSET('" & sourceDataSheet & "'!AJ2, 0, 0, ROWS(AJ:AJ) - ROW(AJ2) + 1),1,FALSE)" 'Secondary 1 TB Period'
		.Range("C26").Formula = "=VLOOKUP(""*"",OFFSET('" & sourceDataSheet & "'!AL2, 0, 0, ROWS(AL:AL) - ROW(AL2) + 1),1,FALSE)" 'Secondary 2 TB Name'
		.Range("C27").Formula = "=VLOOKUP(""*"",OFFSET('" & sourceDataSheet & "'!AM2, 0, 0, ROWS(AM:AM) - ROW(AM2) + 1),1,FALSE)" 'Secondary 2 TB Year'
		.Range("C28").Formula = "=VLOOKUP(""*"",OFFSET('" & sourceDataSheet & "'!AN2, 0, 0, ROWS(AN:AN) - ROW(AN2) + 1),1,FALSE)" 'Secondary 2 TB Period'
	End With
End Sub

Sub CleaningUp()
	'copy the data from Temp to the empty tab (named as the source data sheet)
	'With Worksheets("Temp")
		'.Range("A1", toCell).Copy Worksheets(sourceDataSheet).Range("A1", toCell)
	'End With
	
	
	'Application.DisplayAlerts = False
	'Worksheets("Temp").Delete 'delete the Temp tab
	'Application.DisplayAlerts = True
	
	Dim destinationSheet As String
	
	destinationSheet = "Entity level- PwC"
	DoEvents
	'Worksheets(destinationSheet).Columns("B:G").Calculate
	'RemoveDuplicates destinationSheet, "B33"
	SetStyle destinationSheet, "B33", "K33"
	
	destinationSheet = "Account level- PwC"
	DoEvents
	'Worksheets(destinationSheet).Columns("B:H").Calculate
	'RemoveDuplicates destinationSheet, "B33"
	SetStyle destinationSheet, "B33", "L33"
	
	destinationSheet = "Entity level- Client"
	DoEvents
	'Worksheets(destinationSheet).Columns("B:G").Calculate
	'RemoveDuplicates destinationSheet, "B33"
	SetStyle destinationSheet, "B33", "K33"
	
	destinationSheet = "Account level- Client"
	DoEvents
	'Worksheets(destinationSheet).Columns("B:H").Calculate
	'RemoveDuplicates destinationSheet, "B33"
	SetStyle destinationSheet, "B33", "L33"
	
	OptimiseVBA False 'Resume excel status
End Sub

Sub UpdateColumn(destinationColumn As String, destinationSheet As String, destinationFormula As String, updateSubTotal As Boolean)
	Dim firstRow As String
	Dim lastRow As String
	Dim actualRowNumber As Long
	
	If updateSubTotal = True Then
		With Worksheets(destinationSheet)
			actualRowNumber = .Range("B" & .Rows.Count).End(xlUp).Row - destinationFirstRow + 1
		End With
	Else
		actualRowNumber = sourceDataRowCount
	End If
	
	firstRow = "$" & destinationColumn & "$" & destinationFirstRow 'get first cell to fill the formula
	lastRow = "$" & destinationColumn & "$" & (destinationFirstRow + actualRowNumber - 1) '-1 due to first row is value header
	
	Application.StatusBar = "Updating Column: " & destinationColumn
	
	With Worksheets(destinationSheet)
		.Range(firstRow, lastRow).Formula = destinationFormula
		
		If updateSubTotal = True Then
			Dim subtotalFormula As String
			subtotalFormula = "=SUBTOTAL(9, " & firstRow & ":" & lastRow & ")"
			
			.Range(destinationColumn & (destinationFirstRow - 2)).Formula = subtotalFormula
		End If
	End With
	
End Sub

Sub SetStyle(sheetName As String, startCell As String, endCell As String)
	
	Application.StatusBar = "Updating Styles..."
	
	With Worksheets(sheetName)
		.Range(startCell, endCell).Copy
		
		Dim lastCell As String
		lastCell = .Range(startCell).SpecialCells(xlCellTypeLastCell).Address
		
		With .Range(startCell, lastCell)
			Dim lastRow As Long
			lastRow = .Columns(1).Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
			
			Dim lastColumn As Long
			lastColumn = .Columns.Count
			
			
			With .Range("A" & destinationFirstRow, destinationFirstColumn & destinationFirstRow)
				Dim firstColumnOffset As Long
				firstColumnOffset = .Columns.Count
			End With
			
			lastCell = Cells(lastRow, lastColumn + firstColumnOffset - 1).Address 'get the last cell of the entire table
		End With
		
		.Range(startCell, lastCell).PasteSpecial xlFormats 'paste styles only, no data.
		
	End With
	
	Application.StatusBar = False
	
End Sub

Sub RemoveDuplicates(sheetName As String, fromCell As String)
	With Worksheets(sheetName)
		Dim endCell As String
		
		endCell = .Range(fromCell).SpecialCells(xlCellTypeLastCell).Address 'get last cell
		
		With .Range(fromCell, endCell)
			
			Dim numberOfColumns As Integer
			numberOfColumns = .Columns.Count
			
			Dim columnsArray() As Variant
			Dim i As Long
			ReDim columnsArray(0 To numberOfColumns - 1)
			
			For i = 0 To (numberOfColumns - 1)
				columnsArray(i) = i + 1
				Next i
				
				.RemoveDuplicates Columns:=(columnsArray), Header:=xlNo 'get rid of duplicates
		End With
	End With
End Sub
	
Sub DoPwcEntityLevel()
	'begin to update
	Dim destinationSheet As String
	Dim cellFormula As String
	Dim firstSourceCell As String
	
	destinationSheet = "Entity level- PwC"
	
	'Entity
	firstSourceCell = "B2" 'Source Data Column
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "B", destinationSheet, cellFormula, False
	
	'Entity Name
	firstSourceCell = "X2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "C", destinationSheet, cellFormula, False
	
	'Business Group 2
	firstSourceCell = "Z2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "D", destinationSheet, cellFormula, False
	
	'FS Type
	firstSourceCell = "J2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "E", destinationSheet, cellFormula, False
	
	'Account Type
	firstSourceCell = "K2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "F", destinationSheet, cellFormula, False
	
	'Account Level 2
	firstSourceCell = "O2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "G", destinationSheet, cellFormula, False
	
	Worksheets(destinationSheet).Columns("B:G").Calculate
	RemoveDuplicates destinationSheet, "B33"
	
	'Current Year
	cellFormula = "=IFERROR(SUMIFS('" & sourceDataSheet & "'!H:H, '" & sourceDataSheet & "'!B:B,IF('" & destinationSheet & "'!B33 = 0, """", '" & destinationSheet & "'!B33),'" & sourceDataSheet & "'!O:O,IF('" & destinationSheet & "'!G33 = 0, """", '" & destinationSheet & "'!G33)), 0)" 'define the formula
	UpdateColumn "H", destinationSheet, cellFormula, True
	
	'Comparison
	cellFormula = "=IF('" & destinationSheet & "'!$E33 = ""Balance Sheet"", IFERROR(SUMIFS('" & sourceDataSheet & "'!AK:AK,'" & sourceDataSheet & "'!B:B,IF('" & destinationSheet & "'!B33 = 0, """", '" & destinationSheet & "'!B33),'" & sourceDataSheet & "'!O:O,IF('" & destinationSheet & "'!G33 = 0, """", '" & destinationSheet & "'!G33)),0),IFERROR(SUMIFS('" & sourceDataSheet & "'!AO:AO,'" & sourceDataSheet & "'!B:B,IF('" & destinationSheet & "'!B33 = 0, """", '" & destinationSheet & "'!B33),'" & sourceDataSheet & "'!O:O,IF('" & destinationSheet & "'!G33 = 0, """", '" & destinationSheet & "'!G33)),0))" 'define the formula
	UpdateColumn "I", destinationSheet, cellFormula, True
	
	'Movement
	cellFormula = "=$H33-$I33" 'define the formula
	UpdateColumn "J", destinationSheet, cellFormula, True
	
	'Movement %
	cellFormula = "=IFERROR(ROUND($J33,2)/ROUND($I33,2), 0)" 'define the formula
	UpdateColumn "K", destinationSheet, cellFormula, False
	
	FillHeader destinationSheet
	
End Sub


Sub DoPwcAccountLevel()
	'begin to update
	Dim destinationSheet As String
	Dim cellFormula As String
	Dim firstSourceCell As String
	
	destinationSheet = "Account level- PwC"
	
	'Entity
	firstSourceCell = "B2" 'Source Data Column
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "B", destinationSheet, cellFormula, False
	
	'Entity Name
	firstSourceCell = "X2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "C", destinationSheet, cellFormula, False
	
	'Business Group 2
	firstSourceCell = "Z2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "D", destinationSheet, cellFormula, False
	
	'FS Type
	firstSourceCell = "J2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "E", destinationSheet, cellFormula, False
	
	'Account Type
	firstSourceCell = "K2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "F", destinationSheet, cellFormula, False
	
	'Account Level 2
	firstSourceCell = "O2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "G", destinationSheet, cellFormula, False
	
	'Account
	firstSourceCell = "C2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "H", destinationSheet, cellFormula, False
	
	Worksheets(destinationSheet).Columns("B:H").Calculate
	RemoveDuplicates destinationSheet, "B33"
	
	'Current Year
	firstSourceCell = "H2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "I", destinationSheet, cellFormula, True
	
	'Comparison
	cellFormula = "=IF('" & destinationSheet & "'!$E33 = ""Balance Sheet"", '" & sourceDataSheet & "'!$AK2, '" & sourceDataSheet & "'!$AO2)" 'define the formula
	UpdateColumn "J", destinationSheet, cellFormula, True
	
	'Movement
	cellFormula = "=$I33-$J33" 'define the formula
	UpdateColumn "K", destinationSheet, cellFormula, True
	
	'Movement %
	cellFormula = "=IFERROR(ROUND($K33,2)/ROUND($J33,2), 0)" 'define the formula
	UpdateColumn "L", destinationSheet, cellFormula, False
	
	FillHeader destinationSheet
End Sub


Sub DoClientEntityLevel()
	'begin to update
	Dim destinationSheet As String
	Dim cellFormula As String
	Dim firstSourceCell As String
	
	destinationSheet = "Entity level- Client"
	
	'Entity
	firstSourceCell = "B2" 'Source Data Column
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "B", destinationSheet, cellFormula, False
	
	'Entity Name
	firstSourceCell = "X2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "C", destinationSheet, cellFormula, False
	
	'Business Group 2
	firstSourceCell = "Z2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "D", destinationSheet, cellFormula, False
	
	'FS Type
	firstSourceCell = "J2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "E", destinationSheet, cellFormula, False
	
	'Account Type
	firstSourceCell = "K2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "F", destinationSheet, cellFormula, False
	
	'Account Level 3
	firstSourceCell = "P2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "G", destinationSheet, cellFormula, False
	
	Worksheets(destinationSheet).Columns("B:G").Calculate
	RemoveDuplicates destinationSheet, "B33"
	
	'Current Year
	cellFormula = "=IFERROR(SUMIFS('" & sourceDataSheet & "'!H:H, '" & sourceDataSheet & "'!B:B,IF('" & destinationSheet & "'!B33 = 0, """", '" & destinationSheet & "'!B33),'" & sourceDataSheet & "'!P:P,IF('" & destinationSheet & "'!G33 = 0, """", '" & destinationSheet & "'!G33)), 0)" 'define the formula
	UpdateColumn "H", destinationSheet, cellFormula, True
	
	'Comparison
	cellFormula = "=IF('" & destinationSheet & "'!$E33 = ""Balance Sheet"", IFERROR(SUMIFS('" & sourceDataSheet & "'!AK:AK,'" & sourceDataSheet & "'!B:B,IF('" & destinationSheet & "'!B33 = 0, """", '" & destinationSheet & "'!B33),'" & sourceDataSheet & "'!P:P,IF('" & destinationSheet & "'!G33 = 0, """", '" & destinationSheet & "'!G33)),0),IFERROR(SUMIFS('" & sourceDataSheet & "'!AO:AO,'" & sourceDataSheet & "'!B:B,IF('" & destinationSheet & "'!B33 = 0, """", '" & destinationSheet & "'!B33),'" & sourceDataSheet & "'!P:P,IF('" & destinationSheet & "'!G33 = 0, """", '" & destinationSheet & "'!G33)),0))" 'define the formula
	UpdateColumn "I", destinationSheet, cellFormula, True
	
	'Movement
	cellFormula = "=$H33-$I33" 'define the formula
	UpdateColumn "J", destinationSheet, cellFormula, True
	
	'Movement %
	cellFormula = "=IFERROR(ROUND($J33,2)/ROUND($I33,2), 0)" 'define the formula
	UpdateColumn "K", destinationSheet, cellFormula, False
	
	FillHeader destinationSheet
End Sub


Sub DoClientAccountLevel()
	'begin to update
	Dim destinationSheet As String
	Dim cellFormula As String
	Dim firstSourceCell As String
	
	destinationSheet = "Account level- Client"
	
	'Entity
	firstSourceCell = "B2" 'Source Data Column
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "B", destinationSheet, cellFormula, False
	
	'Entity Name
	firstSourceCell = "X2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "C", destinationSheet, cellFormula, False
	
	'Business Group 2
	firstSourceCell = "Z2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "D", destinationSheet, cellFormula, False
	
	'FS Type
	firstSourceCell = "J2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "E", destinationSheet, cellFormula, False
	
	'Account Type
	firstSourceCell = "K2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "F", destinationSheet, cellFormula, False
	
	'Account Level 3
	firstSourceCell = "P2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "G", destinationSheet, cellFormula, False
	
	'Account
	firstSourceCell = "C2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "H", destinationSheet, cellFormula, False
	
	Worksheets(destinationSheet).Columns("B:H").Calculate
	RemoveDuplicates destinationSheet, "B33"
	
	'Current Year
	firstSourceCell = "H2"
	cellFormula = "='" & sourceDataSheet & "'!$" & firstSourceCell 'define the formula
	UpdateColumn "I", destinationSheet, cellFormula, True
	
	'Comparison
	cellFormula = "=IF('" & destinationSheet & "'!$E33 = ""Balance Sheet"", '" & sourceDataSheet & "'!$AK2, '" & sourceDataSheet & "'!$AO2)" 'define the formula
	UpdateColumn "J", destinationSheet, cellFormula, True
	
	'Movement
	cellFormula = "=$I33-$J33" 'define the formula
	UpdateColumn "K", destinationSheet, cellFormula, True
	
	'Movement %
	cellFormula = "=IFERROR(ROUND($K33,2)/ROUND($J33,2), 0)" 'define the formula
	UpdateColumn "L", destinationSheet, cellFormula, False
	
	FillHeader destinationSheet
	
End Sub
	
