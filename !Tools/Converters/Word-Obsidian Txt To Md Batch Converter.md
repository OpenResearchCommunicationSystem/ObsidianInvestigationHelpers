
```vb
Sub ObsidianTxtToMd()
    Dim MyObj As Object, MySource As Object, file As Variant
    Dim folderPath As String
    Dim dialog As FileDialog

    ' Set up the File Dialog
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)

    If dialog.Show = -1 Then ' if OK is pressed
        folderPath = dialog.SelectedItems(1)

        file = Dir(folderPath & "\*.txt")
        While (file <> "")
           Name folderPath & "\" & file As folderPath & "\" & Replace(file, ".txt", ".md")
           file = Dir
        Wend
    End If
End Sub
```