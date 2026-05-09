Attribute VB_Name = "Module2"
Sub CombineFilesToSheets_FSO()

    Dim FolderPath As String
    Dim fso As Object
    Dim FolderObj As Object
    Dim FileObj As Object
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim SheetName As String
    Dim Counter As Long
    Dim Ext As String

    ' เลือกโฟลเดอร์ที่เก็บไฟล์ Excel
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Please select folder that contains Excel reports"
        
        If .Show <> -1 Then
            MsgBox "ยกเลิกการทำงาน", vbInformation
            Exit Sub
        End If
        
        FolderPath = .SelectedItems(1)
    End With

    ' สร้าง FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set FolderObj = fso.GetFolder(FolderPath)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Counter = 0

    ' วนทุกไฟล์ในโฟลเดอร์ที่เลือก
    For Each FileObj In FolderObj.Files

        Ext = LCase(fso.GetExtensionName(FileObj.Name))

        ' เอาเฉพาะไฟล์ Excel และข้ามไฟล์ชั่วคราว
        If (Ext = "xlsx" Or Ext = "xlsm" Or Ext = "xls") _
            And Left(FileObj.Name, 2) <> "~$" _
            And FileObj.Name <> ThisWorkbook.Name Then

            ' เปิดไฟล์ต้นทาง
            Set wbSource = Workbooks.Open(FileObj.Path)
            Set wsSource = wbSource.Sheets(1)

            ' เพิ่มชีทใหม่ในไฟล์ Macro
            Set wsNew = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))

            ' ตั้งชื่อชีทจากชื่อไฟล์
            SheetName = fso.GetBaseName(FileObj.Name)

            ' ล้างอักขระที่ใช้ตั้งชื่อชีทไม่ได้
            SheetName = Replace(SheetName, "\", "_")
            SheetName = Replace(SheetName, "/", "_")
            SheetName = Replace(SheetName, ":", "_")
            SheetName = Replace(SheetName, "*", "_")
            SheetName = Replace(SheetName, "?", "_")
            SheetName = Replace(SheetName, "[", "_")
            SheetName = Replace(SheetName, "]", "_")

            ' จำกัดชื่อชีทไม่เกิน 31 ตัวอักษร
            SheetName = Left(SheetName, 31)

            ' ตั้งชื่อชีท ถ้าชื่อซ้ำให้ใส่เลขต่อท้าย
            On Error Resume Next
            wsNew.Name = SheetName
            If Err.Number <> 0 Then
                Err.Clear
                wsNew.Name = Left(SheetName, 25) & "_" & Format(Counter + 1, "00")
            End If
            On Error GoTo 0

            ' Copy ข้อมูลจากชีทแรกของไฟล์ต้นทาง
            wsSource.UsedRange.Copy Destination:=wsNew.Range("A1")

            ' ปิดไฟล์ต้นทางโดยไม่ Save
            wbSource.Close SaveChanges:=False

            Counter = Counter + 1

        End If

    Next FileObj

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "รวมไฟล์แยกชีทเสร็จแล้ว จำนวน " & Counter & " ไฟล์", vbInformation

End Sub


Sub FillRestaurantOnly_AllSheets_Buffe()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentRestaurant As String
    
    Set wb = ThisWorkbook
    
    For Each ws In wb.Worksheets
        
        currentRestaurant = ""
        lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        
        For i = 1 To lastRow
            
            ' ถ้าเจอคำว่า "All" ในคอลัมน์ A ให้เก็บชื่อร้าน
            If ws.Cells(i, 1).Value <> "" Then
                If InStr(1, ws.Cells(i, 1).Value, "All", vbTextCompare) > 0 Then
                    currentRestaurant = ws.Cells(i, 1).Value
                End If
            End If
            
            ' ถ้าในคอลัมน์ B มีเลข ให้ใส่ชื่อร้านลงคอลัมน์ H
            If ws.Cells(i, 2).Value <> "" And IsNumeric(ws.Cells(i, 2).Value) Then
                ws.Cells(i, 8).Value = currentRestaurant
            End If
            
        Next i
        
    Next ws
    
    MsgBox "ใส่ชื่อร้านเรียบร้อยทุกชีท!"

End Sub

Sub FillDateToColumnI_AllSheets_Buffe()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim arr() As String
    Dim datePart As String
    Dim dateArr() As String
    Dim sheetDate As Date
    Dim validSheet As Boolean
    
    Set wb = ThisWorkbook
    
    For Each ws In wb.Worksheets
        validSheet = True
        
        arr = Split(ws.Name, "_")
        
        ' ตรวจสอบชื่อชีทมีส่วนมากกว่า 2 (เพื่อแยกส่วนวันที่)
        If UBound(arr) < 2 Then
            validSheet = False
        Else
            datePart = arr(UBound(arr)) ' ส่วนท้ายสุด
            
            On Error GoTo InvalidFormat
            If InStr(datePart, "-") > 0 Then
                ' รูปแบบ dd-mm-yy
                dateArr = Split(datePart, "-")
                If UBound(dateArr) <> 2 Then
                    validSheet = False
                Else
                    sheetDate = DateSerial(2000 + CInt(dateArr(2)), CInt(dateArr(1)), CInt(dateArr(0)))
                End If
            Else
                ' รูปแบบ dd_mm_yy
                If UBound(arr) < 4 Then
                    validSheet = False
                Else
                    sheetDate = DateSerial(2000 + CInt(arr(4)), CInt(arr(3)), CInt(arr(2)))
                End If
            End If
            On Error GoTo 0
        End If
        
        If validSheet Then
            ' หา last row (อิงคอลัม B)
            lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            
            ' ใส่วันที่ลงคอลัม I เฉพาะแถวที่มีข้อมูลในคอลัม B
            For i = 2 To lastRow
                If ws.Cells(i, 2).Value <> "" And IsNumeric(ws.Cells(i, 2).Value) Then
                    ws.Cells(i, 9).Value = sheetDate   ' คอลัม I = 9
                End If
            Next i
            
            ' ตั้งรูปแบบวันที่ให้เป็น dd/mm/yyyy
            ws.Columns(9).NumberFormat = "dd/mm/yyyy"
        End If
        
ContinueLoop:
    Next ws
    
    MsgBox "ใส่วันที่เรียบร้อยทุกชีท!"

    Exit Sub
    
InvalidFormat:
    validSheet = False
    Resume ContinueLoop

End Sub

Sub KeepOnlyNumericRowsInColA_AllSheets_Buffe()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set wb = ThisWorkbook
    
    For Each ws In wb.Worksheets
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' ลบจากล่างขึ้นบน เพื่อไม่พังลำดับ
        For i = lastRow To 1 Step -1
            If Not IsNumeric(ws.Cells(i, 1).Value) Or ws.Cells(i, 1).Value = "" Then
                ws.Rows(i).Delete
            End If
        Next i
        
    Next ws
    
    MsgBox "ลบแถวที่คอลัม A ไม่ใช่ตัวเลขเรียบร้อยทุกชีท!"

End Sub

Sub CombineSheets_AllRows_IN_File()

    Dim ws As Worksheet
    Dim wsSummary As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim pasteRow As Long

    ' ?????/??????? Summary
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    If wsSummary Is Nothing Then
        Set wsSummary = ThisWorkbook.Sheets.Add
        wsSummary.Name = "Summary"
    Else
        wsSummary.Cells.Clear
    End If
    On Error GoTo 0

    pasteRow = 1

    For Each ws In ThisWorkbook.Sheets
        
        If ws.Name <> "Summary" Then
            
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            
            ' ? ????????? ???????????
            ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Copy _
            wsSummary.Cells(pasteRow, 1)
            
            pasteRow = wsSummary.Cells(wsSummary.Rows.Count, 1).End(xlUp).Row + 1
            
        End If
        
    Next ws

    MsgBox "??????????????????!", vbInformation

End Sub
