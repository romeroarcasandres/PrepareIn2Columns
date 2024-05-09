Attribute VB_Name = "Preparein2Columns"
Sub PrepareIn2Columns()
    ' Creates a new Word document
    Dim newDoc As Document
    Set newDoc = Documents.Add
    
    ' Defines the column count for the table
    Dim columnCount As Integer
    columnCount = 2
    
    ' Adds a table to the new document
    Dim newTable As Table
    Set newTable = newDoc.Tables.Add(Range:=Selection.Range, NumRows:=1, NumColumns:=columnCount)
    
    ' Opens a dialog box to select the MS Word Document in .docx or .doc format
    OpenWordDocumentWithDialog
    
    ' Copies content from existing document to the new table
    For Each para In ActiveDocument.Paragraphs
        ' Adds a new row to the table
        newTable.Rows.Add
        
        ' Copies the content to both columns of the new table
        newTable.cell(newTable.Rows.Count, 1).Range.Text = para.Range.Text
        newTable.cell(newTable.Rows.Count, 2).Range.Text = para.Range.Text
    Next para
    
    ' Closes the existing document without saving changes
    ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
    
    ' Deletes empty rows in the new table
    Dim oRow As Row
    Dim cellText As String
    
    ' Assumes the first table in the document
    ' If multiple tables exist, specify which one
    Set newTable = ActiveDocument.Tables(1)
    
    ' Loops backwards to account for changing row count when deleting rows
    For i = newTable.Rows.Count To 1 Step -1
        ' Cleans the cell text by replacing the end-of-cell marker and carriage return
        cellText = Replace(newTable.Rows(i).Cells(1).Range.Text, Chr(7), "")
        ' Also replace carriage return, which is present in empty cells
        cellText = Replace(cellText, Chr(13), "")
        
        ' Checks if the cell text in the first column is empty after cleanup
        If cellText = "" Then
            ' Deletes the entire row
            newTable.Rows(i).Delete
        End If
    Next i
    
    ' Selects the first column
    newTable.Columns(1).Select
    
    ' Hides the selected text
    Selection.Font.Hidden = True

    With Selection
        ' Moves the selection to the end of the document
        .Collapse Direction:=wdCollapseEnd
    End With
    
    ' Displays a message box after completing the actions
    MsgBox "Finished", vbInformation, "Status"
    
    ' Activates the new Word document window
    newDoc.Activate

End Sub

Sub OpenWordDocumentWithDialog()
    ' Variable to store the selected file path
    Dim selectedDocPath As String
    
    ' Creates a FileDialog object as a File Picker dialog box
    With Application.fileDialog(msoFileDialogFilePicker)
        ' Filters to show only Word documents
        .Filters.Clear
        .Filters.Add "Word Documents", "*.docx;*.doc", 1
        
        ' Sets the dialog title
        .Title = "Select a Word Document"
        
        ' Shows the dialog and checks if the user made a selection
        If .Show = -1 Then
            ' Gets the path of the selected file
            selectedDocPath = .SelectedItems(1)
        Else
            ' User canceled the file selection
            MsgBox "No file selected.", vbExclamation, "Canceled"
            Exit Sub
        End If
    End With
    
    ' Opens the selected Word document
    Documents.Open fileName:=selectedDocPath
End Sub

