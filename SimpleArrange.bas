Sub Main
	On Error GoTo ErrorHandler

	Dim PCBDoc
	Dim comp
	Dim i As Long
	Dim currentX As Double
	Dim currentY As Double
	Dim newX As Double
	Dim newY As Double
	Dim deltaX As Double
	Dim deltaY As Double
	Dim baseX As Double
	Dim baseY As Double
	Dim nextX As Double
	Dim padding As Double
	Dim compWidth As Double
	Dim pApp

	Set PCBDoc = ActiveDocument
	If PCBDoc Is Nothing Then
		MsgBox "Please open PCB file first", vbCritical
		Exit Sub
	End If

	padding = 0.2
	baseX = 0
	baseY = 0
	nextX = baseX
	i = 0

	Set pApp = PCBDoc.Application
	pApp.LockServer

	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		With comp.BBox
			currentX = (.Xmin + .Xmax) / 2
			currentY = (.Ymin + .Ymax) / 2
			compWidth = .Xmax - .Xmin
		End With

		If i = 0 Then
			baseX = (.Xmin)
			baseY = (.Ymin)
			nextX = baseX
		End If

		newX = nextX + compWidth / 2
		newY = baseY

		deltaX = newX - currentX
		deltaY = newY - currentY

		comp.Move deltaX, deltaY

		nextX = nextX + compWidth + padding
		i = i + 1
	Next comp

	pApp.UnlockServer

	PCBDoc.Refresh
	MsgBox "Arranged " & i & " components", vbInformation

	Exit Sub
ErrorHandler:
	MsgBox "Error: " & Err.Description, vbCritical
End Sub
