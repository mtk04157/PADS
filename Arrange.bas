Sub Main
	On Error GoTo ErrorHandler

	Dim PCBDoc
	Dim comp
	Dim i As Long
	Dim spacing As Double
	Dim pApp

	Set PCBDoc = ActiveDocument
	If PCBDoc Is Nothing Then
		MsgBox "Please open PCB file first", vbCritical
		Exit Sub
	End If

	spacing = 2.0
	i = 0

	Set pApp = PCBDoc.Application
	pApp.LockServer

	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		On Error Resume Next
		comp.Move i * spacing, 0
		On Error GoTo ErrorHandler
		i = i + 1
	Next comp

	pApp.UnlockServer

	PCBDoc.Refresh
	MsgBox "Arranged " & i & " components", vbInformation

	Exit Sub
ErrorHandler:
	pApp.UnlockServer
	MsgBox "Error at component " & i & ": " & Err.Description, vbCritical
End Sub
