Option Explicit
' ****************
' Macro to insert rows into all sheets
' Hatami Bukhari (Althemier)
' Source: https://spreadsheetvault.com/insert-rows-worksheets/
' There are modification to the original code to use it for delete
' ****************

Sub anotherDeleteRowsSheets()

    ' Disable excel properties before macro runs
    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    ' Decale object variables
    Dim ws As Worksheet, iCountRows As Integer
    Dim activeSheet As Worksheet, activeRow As Long
    Dim startSheet As String

    ' State activeRow
    activeRow = ActiveCell.Row

    ' Save inital active sheet selection
    startSheet = ThisWorkbook.activeSheet.Name

    ' Trigger input message to appear - in terms of how many rows to insert
    iCountRows = Application.InputBox(Prompt:="Enter 1" & activeRow & "?", Type:=1)

    'Error handling - end the macro if a zero, negative integer or non-integer value is entered
    'If iCountRows = False Or iCountRows <= 0 Then End

    ' Loop through the worksheets in active workbook
    For Each ws In ActiveWorkbook.Sheets

        ws.Activate
        Rows(activeRow & ":" & activeRow + iCountRows - 1).EntireRow.Delete

    Next ws

        ' Move cursor back to initial worksheet
        Worksheets(startSheet).Select
        Range("A1").Select

    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
    End With

End Sub



