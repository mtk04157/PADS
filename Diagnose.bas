Sub Main
	Dim PCBDoc
	Dim comp
	Dim count As Long

	Set PCBDoc = ActiveDocument
	count = 0

	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		count = count + 1
	Next comp

	MsgBox "Selected components: " & count, vbInformation
End Sub
