Option Explicit

' =========================================================================
' Purpose: Combine the 1st sheet of all .xls and .xlsx files in a folder
'          into the current workbook as individual sheets.
' Compatibility: 32-bit / 64-bit (Excel 2019 and older)
' Exclusions: Skips .xlsm, .xlsb, and other file types.
' =========================================================================

Sub CombineXlsAndXlsxOnly()
    Dim FolderPath As String    ' Path to the target folder
    Dim FileName As String      ' Name of the current file being processed
    Dim Ext As String           ' Variable to store file extension for filtering
    Dim wbSource As Workbook    ' Reference to the source workbook to be copied
    Dim wsMaster As Workbook    ' Reference to this workbook (the destination)
    Dim FD As FileDialog        ' Folder picker dialog object
    
    Set wsMaster = ThisWorkbook
    
    ' --- Folder Selection Dialog ---
    Set FD = Application.FileDialog(msoFileDialogFolderPicker)
    FD.Title = "Select the folder containing Excel files"
    
    If FD.Show = True Then
        FolderPath = FD.SelectedItems(1) & "\"
    Else
        Exit Sub ' Exit if user cancels the dialog
    End If

    ' Search for files starting with .xls (captures .xls, .xlsx, .xlsm, etc.)
    FileName = Dir(FolderPath & "*.xls*")
    
    ' Performance Optimization: Disable screen flicker, events, and automatic calculation
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    ' Loop through all files found in the folder
    Do While FileName <> ""
        ' Get the file extension and convert to lowercase
        Ext = LCase(Mid(FileName, InStrRev(FileName, ".")))
        
        ' Strict Filter: Only process .xls or .xlsx (excludes .xlsm, etc.)
        If Ext = ".xls" Or Ext = ".xlsx" Then
            ' Skip this workbook if it is located in the target folder
            If FileName <> wsMaster.Name Then
                
                ' Open source file in Read-Only mode for speed and safety
                Set wbSource = Workbooks.Open(FileName:=FolderPath & FileName, ReadOnly:=True)
                
                ' Copy the first sheet to the end of the Master workbook
                wbSource.Sheets(1).Copy After:=wsMaster.Sheets(wsMaster.Sheets.Count)
                
                ' Rename the new sheet to the filename (limit 31 chars for Excel rules)
                On Error Resume Next
                wsMaster.Sheets(wsMaster.Sheets.Count).Name = Left(FileName, 31)
                On Error GoTo 0
                
                ' Close source file without saving changes
                wbSource.Close SaveChanges:=False
            End If
        End If
        
        ' Retrieve the next file name
        FileName = Dir()
    Loop

    ' Restore Application settings
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With

    MsgBox "Process Complete! Only .xls and .xlsx files were imported.", vbInformation
End Sub
