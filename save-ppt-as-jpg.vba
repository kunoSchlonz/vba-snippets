Private Sub CommandButton1_Click()


Dim fd As FileDialog
Dim pfad As String

Set fd = Application.FileDialog(msoFileDialogSaveAs)

With fd

If .Show = -1 Then
MsgBox "Datei wird unter " & .SelectedItems(1) & " gespeichert"

For Each vrtSelectedItem In .SelectedItems
ActivePresentation.SaveAs "C:\test\test.jpg", ppSaveAsJPG
fd.Execute

Next vrtSelectedItem

Else
End If
End With


Set fd = Nothing

ActiveWindow.Selection.SlideRange(1).Export "C:\test\aaa.jpg", "JPG", 4000, 3000

End Sub
