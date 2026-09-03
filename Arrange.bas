Sub Main
	On Error GoTo ErrorHandler

	Dim PCBDoc
	Dim comp
	Dim i As Long
	Dim spacing As Double

	Set PCBDoc = ActiveDocument
	If PCBDoc Is Nothing Then
		MsgBox "Please open PCB file first", vbCritical
		Exit Sub
	End If

	spacing = 2.0
	i = 0

	Dim pApp
	Set pApp = PCBDoc.Application
	pApp.LockServer

	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		comp.Move i * spacing, 0
		i = i + 1
	Next comp

	pApp.UnlockServer

	PCBDoc.Refresh
	MsgBox "Arranged " & i & " components", vbInformation

	Exit Sub
ErrorHandler:
	MsgBox "Error: " & Err.Description, vbCritical
End Sub
