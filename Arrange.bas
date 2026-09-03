Sub Main
	On Error GoTo ErrorHandler

	Dim PCBDoc
	Dim comp
	Dim i As Long
	Dim spacing As Double
	Dim currentX As Double
	Dim currentY As Double
	Dim newX As Double
	Dim newY As Double
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
		currentX = comp.BBox.Xmin + (comp.BBox.Xmax - comp.BBox.Xmin) / 2
		currentY = comp.BBox.Ymin + (comp.BBox.Ymax - comp.BBox.Ymin) / 2
		
		newX = currentX + i * spacing
		newY = currentY

		pApp.ExecuteCommand("Move")
		PCBDoc.MainView.MouseMove(newX, newY)
		pApp.ExecuteCommand("Complete Move", newX, newY)

		i = i + 1
	Next comp

	pApp.UnlockServer

	PCBDoc.Refresh
	MsgBox "Arranged " & i & " components", vbInformation

	Exit Sub
ErrorHandler:
	pApp.UnlockServer
	MsgBox "Error: " & Err.Description, vbCritical
End Sub
