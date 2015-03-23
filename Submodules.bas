' To access original Personal file: C:\Users\dlachasse\AppData\Roaming\Microsoft\Excel\XLSTART\Bak

''''''''''''''''''''''''''''''''''
'''''''''''''' SUBS ''''''''''''''
''''''''''''''''''''''''''''''''''

Public latestFile As String

Sub OpenMostRecentFile(path As String)

    Dim file As String
    Dim latestDate, LMD As Date

    file = Dir(path & "*.txt", vbNormal)

    Do While file <> Empty
        LMD = FileDateTime(path & file)
        If LMD > latestDate Then
            latestFile = file
            latestDate = LMD
        End If

        file = Dir
    Loop

    Workbooks.Open path & latestFile

End Sub

Sub PrepInventorySync()

Dim pad As Integer
Dim cell As Range
Dim filePre As String

OpenMostRecentFile ("\\faserv\ds_Supplier\Visr\Output\")

'' Completion message
    Dim Msg, Style, Title, MyString
    Msg = "Has Visr utility completed processing and " & latestFile & " the correct output file?"   ' Define message."
    Style = vbYesNo + vbSystemModal + vbDefaultButton1   ' Define buttons.
    Title = "Continue Prompt"   ' Define title.
    
          ' Display message.
        response = MsgBox(Msg, Style, Title)
        If response = vbYes Then   ' User chose Yes.
           GoTo ProcessFile
        Else    ' User chose No.
            ActiveWorkbook.Close SaveChanges:=False
            Exit Sub
        End If

ProcessFile:
    filePre = Format(Date, "yyyymmdd") & ".txt"
    
    '   Format the output of the Visr utility
    Columns(1).Delete Shift:=xlToLeft
    Columns(2).Insert Shift:=xlToRight
    Cells(1, 2).Value = "Price"
    
    With ActiveSheet
        .Range("A:C").RemoveDuplicates Columns:=1, Header:=xlYes ' Removes duplicate skus
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("C" & LastRow + 1).End(xlDown).ClearContents ' Removes excess zeroes from duplicate skus
    End With
    
    pad = InputBox("How much do you want to pad inventory quantities?")

    '   Cycle through cells and remove padding amount
    For Each cell In Range("C2", "C" & LastRow)
        If cell.Value <= pad Then
            cell.Value = 0
        ElseIf cell.Value < 0 Then
            cell.Value = 0
        Else
            cell.Value = cell.Value - pad
        End If
        
        Next cell

    '   Save file for uploading
    ChDir "\\faserv\ds_Supplier\Visr\Upload"
    ActiveWorkbook.SaveAs Filename:= _
        "\\faserv\ds_Supplier\Visr\Upload\" & filePre, FileFormat:=xlText, CreateBackup:=False
    Workbooks(filePre).Close SaveChanges:=False
    
    MsgBox (filePre & " successfully formatted!"), vbOKOnly

End Sub

Sub SSInventory()

Dim LR As Long
Dim cell As Range
Dim sku, AliasSKU As String
Dim iVal As Integer

    OpenMostRecentFile ("P:\Amazon\SSInventory\")

    Dim Msg, Style, Title, MyString
    Msg = "Is " & latestFile & " the correct file?"
    Style = vbYesNo + vbSystemModal + vbDefaultButton1
    Title = "Continue Prompt"
    
        response = MsgBox(Msg, Style, Title)
        If response = vbYes Then
           GoTo ProcessSS
        Else
            ActiveWorkbook.Close SaveChanges:=False
            Exit Sub
        End If

ProcessSS:
    Application.DisplayAlerts = False

    Sheets.Add After:=Sheets(Sheets.Count)
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array(Array( _
        "ODBC;DSN=SE Data;Description=StoneEdge data connection;UID=dlachasse;Trusted_Connection=Yes;APP=2007 Microsoft Office system;WSID=FA" _
        ), Array( _
        "0015;DATABASE=SE Data;LANGUAGE=us_english;Network=DBMSSOCN;Address=faserv\sqlexpress,1433" _
        )), Destination:=Range("$A$1")).QueryTable
        .CommandText = Array( _
        "SELECT AliasSKUs.ParentSKU, AliasSKUs.AliasSKU  FROM ""SE Data"".dbo.AliasSKUs AliasSKUs, ""SE Data"".dbo.Inventory Inv" _
        , _
        "entory  WHERE AliasSKUs.ParentSKU = Inventory.LocalSKU AND ((Inventory.Category='Licensed Shirts'))" _
        )
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .SourceConnectionFile = _
        "C:\Users\dlachasse\AppData\Roaming\Microsoft\Queries\LicensedShirtAliasSKUs.dqy"
        .ListObject.DisplayName = "Table_LicensedShirtAliasSKUs"
        .Refresh BackgroundQuery:=False
    End With

    Sheets(1).Select
    For x = 1 To LR

    For Each cell In Range("A2", "A" & LR)
        
        sku = cell.Value
        iVal = Application.WorksheetFunction.CountIf(Sheets(2).Range("A:A"), sku)
        If iVal = 0 Then GoTo skip
        i = 0
        Do Until i > iVal
            If i > iVal Then GoTo skip
            Select Case i
                Case 0
                    Set newSku = Sheets(2).Columns(1).Cells.Find(sku)
                    AliasSKU = newSku.Offset(0, 1).Value
                    cell.Value = AliasSKU
                Case 1
                    Set lastSKU = Sheets(2).Columns(1).Cells.Find(sku)
                    Set newSku = Sheets(2).Columns(1).Cells.FindNext(lastSKU)
                    AliasSKU = newSku.Offset(0, 1).Value
                    cell.Offset(1, 0).EntireRow.Insert
                    cell.Offset(1, 0).Value = AliasSKU
                    cell.Offset(1, 2).Value = cell.Offset(0, 2).Value
                Case 2
                    Set lastSKU = Sheets(2).Columns(1).Cells.Find(sku)
                    Set newSku = Sheets(2).Columns(1).Cells.FindNext(lastSKU)
                    Set newSku = Sheets(2).Columns(1).Cells.FindNext(newSku)
                    AliasSKU = newSku.Offset(0, 1).Value
                    cell.Offset(2, 0).EntireRow.Insert
                    cell.Offset(2, 0).Value = AliasSKU
                    cell.Offset(2, 2).Value = cell.Offset(0, 2).Value
                Case 3
                    Set lastSKU = Sheets(2).Columns(1).Cells.Find(sku)
                    Set newSku = Sheets(2).Columns(1).Cells.FindNext(lastSKU)
                    Set newSku = Sheets(2).Columns(1).Cells.FindNext(newSku)
                    Set newSku = Sheets(2).Columns(1).Cells.FindNext(newSku)
                    AliasSKU = newSku.Offset(0, 1).Value
                    cell.Offset(3, 0).EntireRow.Insert
                    cell.Offset(3, 0).Value = AliasSKU
                    cell.Offset(3, 2).Value = cell.Offset(0, 2).Value
            End Select
            i = i + 1
        Loop
skip:
        x = x + 1
        Application.StatusBar = "Progress: " & x & " of " & LR & ": " & Format(x / LR, "0%")
        Next cell
    Next x

    Sheets(1).Range("A:C").RemoveDuplicates Columns:=1, Header:=xlYes
    Application.StatusBar = False

    Sheets(2).Delete
    
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(1).Range("D:D").Copy
    Sheets(2).Range("A:A").Insert
    Sheets(1).Select
    ActiveSheet.Range("D:D").ClearContents

    filePre = Format(Date, "yyyymmdd")
    ChDir "\\faserv\ds_Supplier\Upload"
    ActiveWorkbook.SaveAs Filename:="\\faserv\ds_Supplier\Upload\" & filePre & " SS UK", FileFormat:=xlText, CreateBackup:=False

    Sheets(2).Range("A:A").Copy
    Sheets(1).Range("D:D").Insert
    Sheets(2).Delete
    Sheets(1).Select
    With ActiveSheet
        .Range("A:D").AutoFilter Field:=4, Criteria1:="Yes"
        Range(Rows("2:2"), Selection.End(xlDown)).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.Delete Shift:=xlUp
        .AutoFilterMode = False
        Columns(4).ClearContents
        LR = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With

    ChDir "\\faserv\ds_Supplier\Upload"
    ActiveWorkbook.SaveAs Filename:="\\faserv\ds_Supplier\Upload\" & filePre & " SS", FileFormat:=xlText, CreateBackup:=False
    ActiveWorkbook.Close False

    Application.DisplayAlerts = True

End Sub

Sub ProfitBuilder()
    Dim cart As String

    ActiveCell.Offset(-1, 0).Value = "Price + Shipping"
    ActiveCell.Offset(-1, 1).Value = "Amazon commission"
    ActiveCell.Offset(-1, 2).Value = "Avg Ship Cost"
    ActiveCell.Offset(-1, 3).Value = "Profit"
    cart = Range("N14").Value

    Do While ActiveCell.Offset(0, -1).Value <> ""
        
        ActiveCell.FormulaR1C1 = "=NetProfitBreakdown(""" & cart & """,RC[-1],RC[-2],RC[-3],1)"
        ActiveCell.Offset(0, 1).FormulaR1C1 = "=NetProfitBreakdown(""" & cart & """,RC[-2],RC[-3],RC[-4],2)"
        ActiveCell.Offset(0, 2).FormulaR1C1 = "=NetProfitBreakdown(""" & cart & """,RC[-3],RC[-4],RC[-5],3)"
        ActiveCell.Offset(0, 3).FormulaR1C1 = "=NetProfitBreakdown(""" & cart & """,RC[-4],RC[-5],RC[-6],4)"

        ActiveCell.Offset(1, 0).Select

    Loop

End Sub

Sub prepInvSync()

Dim cart, pad As Integer
Dim LastRow, x As Long
Dim cell, newSku As Range
Dim filePre, sku, AliasSKU, suf As String

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .EnableEvents = False
End With

OpenMostRecentFile ("\\faserv\ds_Supplier\Visr\Output\")

    Dim Msg, Style, Title, MyString
    Msg = "Has supplier utility completed processing and " & latestFile & " the correct output file?"
    Style = vbYesNo + vbSystemModal + vbDefaultButton1   ' Define buttons.
    Title = "Continue Prompt"   ' Define title.
    
          ' Display message.
        response = MsgBox(Msg, Style, Title)
        If response = vbYes Then   ' User chose Yes.
           GoTo ProcessFile
        Else    ' User chose No.
            ActiveWorkbook.Close SaveChanges:=False
            Exit Sub
        End If

ProcessFile:
    cart = InputBox("Which cart number are you processing? /n 1 = Blank Apparel 4 = HiveTeesUK", "Cart ID")
    
    '   Format the output of the Visr utility
    Columns(1).Delete Shift:=xlToLeft
    Columns(2).Insert Shift:=xlToRight
    Cells(1, 2).Value = "Price"
    
    With ActiveSheet
        .Range("A:C").RemoveDuplicates Columns:=1, Header:=xlYes ' Removes duplicate skus
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("C" & LastRow + 1).End(xlDown).ClearContents ' Removes excess zeroes from duplicate skus
    End With
    
    pad = InputBox("How much do you want to pad inventory quantities?")

    '   Cycle through cells and remove padding amount
    For Each cell In Range("C2", "C" & LastRow)
        If cell.Value <= pad Then
            cell.Value = 0
        ElseIf cell.Value < 0 Then
            cell.Value = 0
        Else
            cell.Value = cell.Value - pad
        End If
        
        Next cell

    '   Import AliasSKUs records for cart previously specified
    Worksheets.Add
    Sheets(1).Select
    With Sheets(1).ListObjects.Add(SourceType:=0, Source:=Array(Array( _
        "ODBC;DSN=SE Data;Description=StoneEdge data connection;UID=dlachasse;Trusted_Connection=Yes;APP=2007 Microsoft Office system;WSID=FA" _
        ), Array( _
        "0015;DATABASE=SE Data;LANGUAGE=us_english;Network=DBMSSOCN;Address=faserv\sqlexpress,1433" _
        )), Destination:=Range("$A$1")).QueryTable
        .CommandText = Array( _
        "SELECT AliasSKUs.ParentSKU, AliasSKUs.AliasSKU, AliasSKUs.CartID" & Chr(13) & "" & Chr(10) & "FROM ""SE Data"".dbo.AliasSKUs AliasSKUs" & Chr(13) & "" & Chr(10) & "WHERE (AliasSKUs.CartID=" & cart & ")" & Chr(13) & "" & Chr(10) & "ORDER BY AliasSKUs.ParentSKU" _
        )
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Table_Query_from_SE_Data"
        .Refresh BackgroundQuery:=False
    End With
    
    Sheets(2).Select
    Range("A2").Select
    
    '   Status bar progress
    For x = 1 To LastRow

        For Each cell In Range("A2", "A" & LastRow)
        
            sku = cell.Value
            Set newSku = Sheets(1).Columns(1).Cells.Find(sku)
            If newSku Is Nothing Then GoTo skip
            AliasSKU = newSku.Offset(0, 1).Value
            cell.Value = AliasSKU
skip:
            x = x + 1
            Application.StatusBar = "Progress: " & x & " of " & LastRow & ": " & Format(x / LastRow, "0%")
            Next cell
    Next x
    
    Application.StatusBar = False
    
    Sheets(1).Delete
    Select Case cart
        Case 1
            suf = " US"
        Case 4
            suf = " UK"
        Case Else
            suf = ""
    End Select

    '   Save file for uploading
    filePre = Format(Date, "yyyymmdd")
    ChDir "\\faserv\ds_Supplier\Upload"
    ActiveWorkbook.SaveAs Filename:= _
        "\\faserv\ds_Supplier\Upload\" & filePre & suf, FileFormat:=xlText, CreateBackup:=False
    Workbooks(filePre & suf).Close SaveChanges:=False
    
    MsgBox (filePre & " successfully formatted!"), vbOKOnly

End Sub

Sub UpdateImprints()
'
' Uses latest Imprints XML pull to configure SEOM inventory data refresh file
'
Workbooks.Open ("\\faserv\ds_Supplier\Imprints\Process\ImprintsInventoryXMLBare.xlsx")

Dim LR As Long

    ActiveWorkbook.XmlMaps("Imprints_Map").DataBinding.Refresh
    Sheets("DB").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False

    With Sheets("Data")
    
        LR = .Cells(.Rows.Count, "A").End(xlUp).Row

    End With

    Sheets("SE_Import").Select
    Range("A2").FormulaR1C1 = "=INDEX(DB!C[1]:C[2],MATCH(Table1[[#This Row],[item-number]],DB!C[1],0),2)"
    Range("B2").FormulaR1C1 = "=Table1[[#This Row],[price3]]"
    Range("C2").FormulaR1C1 = "=ArrayAdd(Table1[[#This Row],[qty]])"
    Range("A2:C2").AutoFill Destination:=Range("A2:C" & LR)
    Sheets("SE_Import").Calculate
    Cells.Copy
    Cells.PasteSpecial xlPasteValues

    For Each cell In Range("A2:A" & LR)
        If IsError(cell) Then
            cell.EntireRow.Delete
        End If
    Next cell

    Sheets("SE_Import").Range("A:C").RemoveDuplicates Columns:=1, Header:=xlYes


'   Save file for uploading
    filePre = Format(Date, "yyyymmdd")
    ChDir "C:\Users\dlachasse\Desktop\Projects\SEOM\"
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\dlachasse\Desktop\Projects\SEOM\" & filePre & " Imprints", FileFormat:=xlText, CreateBackup:=False
    ActiveWorkbook.Close SaveChanges:=False
    
    MsgBox (filePre & " Imprints" & " successfully formatted!"), vbOKOnly
End Sub

Sub AddParentRow()

Do While ActiveCell <> ""

    If ActiveCell.Value <> ActiveCell.Offset(1, 0) Then
        ActiveCell.Offset(1, 0).EntireRow.Insert
        ActiveCell.Offset(1, 0).Select
    End If

    ActiveCell.Offset(1, 0).Select

Loop

End Sub

Sub QuoteCommaExport()
   ' Dimension all variables.
   Dim DestFile As String
   Dim FileNum As Integer
   Dim ColumnCount As Integer
   Dim RowCount As Integer

   ' Prompt user for destination file name.
   DestFile = InputBox("Enter the destination filename" _
      & Chr(10) & "(with complete path):", "Quote-Comma Exporter")

   ' Obtain next free file handle number.
   FileNum = FreeFile()

   ' Turn error checking off.
   On Error Resume Next

   ' Attempt to open destination file for output.
   Open DestFile For Output As #FileNum

   ' If an error occurs report it and end.
   If Err <> 0 Then
      MsgBox "Cannot open filename " & DestFile
      End
   End If

   ' Turn error checking on.
   On Error GoTo 0

   ' Loop for each row in selection.
   For RowCount = 1 To Selection.Rows.Count

      ' Loop for each column in selection.
      For ColumnCount = 1 To Selection.Columns.Count

         ' Write current cell's text to file with quotation marks.
         Print #FileNum, """" & Selection.Cells(RowCount, _
            ColumnCount).Text & """";

         ' Check if cell is in last column.
         If ColumnCount = Selection.Columns.Count Then
            ' If so, then write a blank line.
            Print #FileNum,
         Else
            ' Otherwise, write a comma.
            Print #FileNum, ",";
         End If
      ' Start next iteration of ColumnCount loop.
      Next ColumnCount
   ' Start next iteration of RowCount loop.
   Next RowCount

   ' Close destination file.
   Close #FileNum
End Sub

