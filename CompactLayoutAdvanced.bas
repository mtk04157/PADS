Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFFFFFF
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Sub Main
	On Error GoTo ErrorHandler

	Dim PCBDoc
	Dim comp
	Dim compArray()
	Dim compNames() As String
	Dim compCount As Long
	Dim i As Long
	Dim pApp
	Dim padding As Double
	Dim minX As Double, minY As Double
	Dim currentX As Double, currentY As Double
	Dim newX As Double, newY As Double
	Dim deltaX As Double, deltaY As Double
	Dim baseX As Double, baseY As Double
	Dim colWidth As Double
	Dim compWidth As Double, compHeight As Double

	Set PCBDoc = ActiveDocument
	If PCBDoc Is Nothing Then
		MsgBox "Please open PCB file first", vbCritical
		Exit Sub
	End If

	padding = 0.2
	compCount = 0
	minX = 999999
	minY = 999999
	colWidth = 0

	Set pApp = PCBDoc.Application
	pApp.LockServer

	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		ReDim Preserve compArray(compCount)
		ReDim Preserve compNames(compCount)
		Set compArray(compCount) = comp
		compNames(compCount) = comp.Name

		With comp.BBox
			compWidth = .Xmax - .Xmin
			compHeight = .Ymax - .Ymin
			currentX = (.Xmin + .Xmax) / 2
			currentY = (.Ymin + .Ymax) / 2

			If currentX < minX Then minX = currentX
			If currentY < minY Then minY = currentY
			If compWidth > colWidth Then colWidth = compWidth
		End With

		compCount = compCount + 1
	Next comp

	pApp.UnlockServer

	If compCount = 0 Then
		MsgBox "No components selected", vbInformation
		Exit Sub
	End If

	baseX = minX
	baseY = minY
	pApp.LockServer

	For i = 0 To compCount - 1
		Set comp = compArray(i)

		With comp.BBox
			currentX = (.Xmin + .Xmax) / 2
			currentY = (.Ymin + .Ymax) / 2
			compWidth = .Xmax - .Xmin
			compHeight = .Ymax - .Ymin
		End With

		newX = baseX + (i * (colWidth + padding)) + (colWidth / 2)
		newY = baseY

		deltaX = newX - currentX
		deltaY = newY - currentY

		comp.Move deltaX, deltaY
	Next i

	pApp.UnlockServer

	PCBDoc.Refresh
	MsgBox "Arranged " & compCount & " components with " & padding & "mm pad-to-pad target spacing", vbInformation

	Exit Sub
ErrorHandler:
	MsgBox "Error: " & Err.Description, vbCritical
End Sub
