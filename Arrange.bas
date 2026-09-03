Sub Main
	On Error GoTo ErrorHandler

	Dim PCBDoc
	Dim comp
	Dim i As Long
	Dim spacing As Double
	Dim pApp
	Dim compCount As Long

	Set PCBDoc = ActiveDocument
	If PCBDoc Is Nothing Then
		MsgBox "Please open PCB file first", vbCritical
		Exit Sub
	End If

	spacing = 2.0
	compCount = 0

	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		compCount = compCount + 1
	Next comp

	If compCount = 0 Then
		MsgBox "No components selected", vbInformation
		Exit Sub
	End If

	Set pApp = PCBDoc.Application
	pApp.LockServer

	i = 0
	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		comp.Move i * spacing, 0
		i = i + 1
	Next comp

	pApp.UnlockServer

	PCBDoc.Refresh
	MsgBox "Arranged " & compCount & " components", vbInformation

	Exit Sub
ErrorHandler:
	pApp.UnlockServer
	MsgBox "Error: " & Err.Description, vbCritical
End Sub
