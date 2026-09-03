Sub Main
	On Error GoTo ErrorHandler

	Dim PCBDoc
	Dim components
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
	Dim spacing As Double
	Dim compWidth As Double
	Dim compHeight As Double
	Dim pApp
	Dim refName As String

	Set PCBDoc = ActiveDocument
	If PCBDoc Is Nothing Then
		MsgBox "Please open PCB file first", vbCritical
		Exit Sub
	End If

	Set pApp = PCBDoc.Application
	pApp.LockServer

	baseX = 0
	baseY = 0
	nextX = baseX
	i = 0

	Set components = PCBDoc.Components

	For Each comp In components
		If comp.Selected Then
			With comp.BBox
				currentX = (.Xmin + .Xmax) / 2
				currentY = (.Ymin + .Ymax) / 2
				compWidth = .Xmax - .Xmin
				compHeight = .Ymax - .Ymin

				If i = 0 Then
					baseX = .Xmin
					baseY = .Ymin
					nextX = baseX
				End If
			End With

			refName = comp.RefDes

			If refName = "U6502" Or refName = "U6503" Then
				spacing = 0
			Else
				spacing = 0.2
			End If

			newX = nextX + compWidth / 2
			newY = baseY + compHeight / 2

			deltaX = newX - currentX
			deltaY = newY - currentY

			comp.Move deltaX, deltaY

			nextX = nextX + compWidth + spacing
			i = i + 1
		End If
	Next comp

	pApp.UnlockServer

	PCBDoc.Refresh
	MsgBox "Arranged " & i & " components", vbInformation

	Exit Sub
ErrorHandler:
	If Not pApp Is Nothing Then
		pApp.UnlockServer
	End If
	MsgBox "Error: " & Err.Description, vbCritical
End Sub
