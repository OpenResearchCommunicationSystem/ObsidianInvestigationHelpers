
```vb
Sub ObsidianDocxToMd()
    Dim xIndex As Long
    Dim xFolder As Variant
    Dim xFileStr As String
    Dim xFilePath As String
    Dim xDlg As FileDialog
    Dim xActPath As String
    Dim xDoc As Document

    Application.ScreenUpdating = False
    Set xDlg = Application.FileDialog(msoFileDialogFolderPicker)
    If xDlg.Show <> -1 Then Exit Sub
    xFolder = xDlg.SelectedItems(1)
    xFileStr = Dir(xFolder & \\*.docx)
    xActPath = ActiveDocument.Path

    While xFileStr <> ""
        xFilePath = xFolder & "\\" & xFileStr
        If xFilePath <> xActPath Then
            Set xDoc = Documents.Open(xFilePath, AddToRecentFiles:=False, Visible:=False)
            xIndex = InStrRev(xFilePath, ".")
            xDoc.SaveAs Left(xFilePath, xIndex - 1) & ".md", FileFormat:=wdFormatText, AddToRecentFiles:=False
            xDoc.Close True
        End If
        xFileStr = Dir()
    Wend
    Application.ScreenUpdating = True

End Sub

```