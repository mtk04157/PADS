Sub Main
	Dim PCBDoc
	Dim comp
	Dim i As Long

	Set PCBDoc = ActiveDocument

	i = 0
	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		comp.Move i * 2, 0
		i = i + 1
	Next comp

	PCBDoc.Refresh
	MsgBox "Done: " & i, vbInformation
End Sub
