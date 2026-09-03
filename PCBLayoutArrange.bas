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
	Dim compCount As Integer
	Dim i As Integer
	Dim minX As Double, minY As Double
	Dim maxX As Double, maxY As Double
	Dim currentX As Double, currentY As Double
	Dim newX As Double, newY As Double
	Dim deltaX As Double, deltaY As Double
	Dim padding As Double
	Dim pApp

	Set PCBDoc = ActiveDocument
	If PCBDoc Is Nothing Then
		MsgBox "Please open PCB file first", vbCritical
		Exit Sub
	End If

	padding = 0.2
	minX = 999999
	minY = 999999
	maxX = -999999
	maxY = -999999
	compCount = 0

	Set pApp = PCBDoc.Application
	pApp.LockServer

	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		ReDim Preserve compArray(compCount)
		Set compArray(compCount) = comp
		compCount = compCount + 1
	Next comp

	pApp.UnlockServer

	If compCount = 0 Then
		MsgBox "No components selected", vbInformation
		Exit Sub
	End If

	For i = 0 To compCount - 1
		Set comp = compArray(i)
		With comp.BBox
			currentX = (.Xmin + .Xmax) / 2
			currentY = (.Ymin + .Ymax) / 2
			If currentX < minX Then minX = currentX
			If currentY < minY Then minY = currentY
		End With
	Next i

	pApp.LockServer

	For i = 0 To compCount - 1
		Set comp = compArray(i)
		With comp.BBox
			currentX = (.Xmin + .Xmax) / 2
			currentY = (.Ymin + .Ymax) / 2
			newX = minX + (i * padding)
			newY = minY
			deltaX = newX - currentX
			deltaY = newY - currentY
			comp.Move deltaX, deltaY
		End With
	Next i

	pApp.UnlockServer

	PCBDoc.Refresh
	MsgBox "Arranged " & compCount & " components in a row with " & padding & "mm spacing", vbInformation

	Exit Sub
ErrorHandler:
	MsgBox "Error: " & Err.Description, vbCritical
End Sub
