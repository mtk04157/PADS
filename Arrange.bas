Sub Main
	On Error GoTo ErrorHandler

	Dim PCBDoc
	Dim comp
	Dim i As Long
	Dim baseX As Double
	Dim baseY As Double
	Dim pApp

	Set PCBDoc = ActiveDocument
	If PCBDoc Is Nothing Then
		MsgBox "Please open PCB file first", vbCritical
		Exit Sub
	End If

	baseX = 0
	baseY = 0
	i = 0

	Set pApp = PCBDoc.Application

	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		pApp.ExecuteCommand("Move")
		PCBDoc.MainView.MouseMove(baseX + i * 2.5, baseY)
		PCBDoc.MainView.MouseMove(baseX + i * 2.5, baseY)
		pApp.ExecuteCommand("Complete Move", baseX + i * 2.5, baseY)
		i = i + 1
	Next comp

	PCBDoc.Refresh
	MsgBox "Arranged " & i & " components", vbInformation

	Exit Sub
ErrorHandler:
	MsgBox "Error: " & Err.Description, vbCritical
End Sub
