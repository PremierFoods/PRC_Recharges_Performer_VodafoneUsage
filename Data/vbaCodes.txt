Sub ExtendFormula(worksheet as String, lastRow as Long)
    Dim formulaCell As Range

    ' Set the worksheet (change "Sheet1" to your sheet name)
    Set ws = ThisWorkbook.Sheets(worksheet)

    if ws.Name = "Summary Data" Then
 	' Set the cell that contains the initial formula
    	Set formulaCell = ws.Range("A4")
	
    	' Extend the formula down to the last row
    	formulaCell.AutoFill Destination:=ws.Range("A4:A" & lastRow)

	 ' Set the cell that contains the initial formula
    	Set formulaCell = ws.Range("B3")
	
    	' Extend the formula down to the last row
    	formulaCell.AutoFill Destination:=ws.Range("B3:B" & lastRow)

    	' Set the cell that contains the initial formula
    	Set formulaCell = ws.Range("C3")

   	 ' Extend the formula down to the last row
    	formulaCell.AutoFill Destination:=ws.Range("C3:C" & lastRow)

    	' Set the cell that contains the initial formula
    	Set formulaCell = ws.Range("D3")

    	' Extend the formula down to the last row
    	formulaCell.AutoFill Destination:=ws.Range("D3:D" & lastRow)

	 ' Set the cell that contains the initial formula
    	Set formulaCell = ws.Range("E3")
	
    	' Extend the formula down to the last row
    	formulaCell.AutoFill Destination:=ws.Range("E3:E" & lastRow)

    	' Set the cell that contains the initial formula
    	Set formulaCell = ws.Range("F3")

   	 ' Extend the formula down to the last row
    	formulaCell.AutoFill Destination:=ws.Range("F3:F" & lastRow)

    	' Set the cell that contains the initial formula
    	Set formulaCell = ws.Range("G3")

    	' Extend the formula down to the last row
    	formulaCell.AutoFill Destination:=ws.Range("G3:G" & lastRow)
    End if
    
    if ws.Name = "Data for journal" then
	' Set the cell that contains the initial formula
    	Set formulaCell = ws.Range("D7")

    	' Extend the formula down to the last row
    	formulaCell.AutoFill Destination:=ws.Range("D7:D" & lastRow)
    End if
  
    if ws.Name = "Journal" Then

    	' Set the cell that contains the initial formula
    	Set formulaCell = ws.Range("B3")

    	' Extend the formula down to the last row
    	formulaCell.AutoFill Destination:=ws.Range("B3:B" & lastRow)
 
       ' Set the cell that contains the initial formula
    	Set formulaCell = ws.Range("E3")

    	' Extend the formula down to the last row
    	formulaCell.AutoFill Destination:=ws.Range("E3:E" & lastRow)


    	' Set the cell that contains the initial formula
    	Set formulaCell = ws.Range("I3")

    	' Extend the formula down to the last row
    	formulaCell.AutoFill Destination:=ws.Range("I3:I" & lastRow)

    	' Set the cell that contains the initial formula
    	Set formulaCell = ws.Range("L3")

   	 ' Extend the formula down to the last row
    	formulaCell.AutoFill Destination:=ws.Range("L3:L" & lastRow)

    	' Set the cell that contains the initial formula
    	Set formulaCell = ws.Range("S3")

    	' Extend the formula down to the last row
    	formulaCell.AutoFill Destination:=ws.Range("S3:S" & lastRow)
    End if

End Sub

Sub CopyColumnDataInRange(worksheet as String, sourceColumn as string, targetColumn as string)
    Dim ws As Worksheet
    Dim headerRow As Range
    Dim sourceCol As Long
    Dim targetCol As Long
    Dim lastRow As Long
    Dim i As Long

    ' Set worksheet and header names
    Set ws = ThisWorkbook.Sheets(worksheet)
    
    ' Set the header row
    Set headerRow = ws.Rows(6)

    ' Find the source and target columns by their header names
    sourceCol = Application.Match(sourceColumn, headerRow, 0)
    targetCol = Application.Match(targetColumn, headerRow, 0)


    ' Find the last row with data in the source column
    lastRow = ws.Cells(ws.Rows.Count, sourceCol).End(xlUp).Row

    ' Loop through each row, starting from row 7 (to skip headers)
    For i = 7 To lastRow
        ws.Cells(i, targetCol).Value = ws.Cells(i, sourceCol).Value
    Next i

End Sub

Sub RemoveDuplicatesInColumn(worksheet as string)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(worksheet) ' Change "Sheet1" to your sheet name

    ' Define the range you want to check for duplicates (e.g., Column A)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row ' Find the last used row in Column A

    ' Remove duplicates in Column A
    ws.Range("C7:C" & lastRow).RemoveDuplicates Columns:=1, Header:=xlYes
End Sub

Sub ExcludeSpecificValueInPivot(worksheet as string, excludeList as string, pivotName as string, pivotField as string)
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim pi As PivotItem

    ' Set worksheet and pivot table
    Set ws = ThisWorkbook.Worksheets(worksheet)
    Set pt = ws.PivotTables(pivotName)
    Set pf = pt.PivotFields(pivotField)

'MsgBox "Field '" & pf & "' found in PivotTable '" & pt & "'."

    ' Loop through each item in the pivot field
	
    For Each pi In pf.PivotItems

		if pi.Visible= False then
		  pi.Visible = True
               End If
        For j = LBound(split(excludeList,",")) To UBound(split(excludeList,","))
            If pi.Name = split(excludeList,",")(j) Then
		'MsgBox "Exclude '" & pi.Name & "' from PivotTable '" & pt & "'."
                pi.Visible = False
                Exit For
            End If
        Next j
    Next pi
End Sub
