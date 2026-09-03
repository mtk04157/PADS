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
	Dim compWidth()
	Dim compHeight()
	Dim compCount As Long
	Dim i As Long, j As Long
	Dim rows As Long, cols As Long
	Dim minX As Double, minY As Double
	Dim currentX As Double, currentY As Double
	Dim newX As Double, newY As Double
	Dim deltaX As Double, deltaY As Double
	Dim padding As Double
	Dim pApp
	Dim rowHeight As Double
	Dim colWidth As Double
	Dim baseX As Double, baseY As Double

	Set PCBDoc = ActiveDocument
	If PCBDoc Is Nothing Then
		MsgBox "Please open PCB file first", vbCritical
		Exit Sub
	End If

	padding = 0.2
	minX = 999999
	minY = 999999
	compCount = 0

	Set pApp = PCBDoc.Application
	pApp.LockServer

	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		ReDim Preserve compArray(compCount)
		ReDim Preserve compWidth(compCount)
		ReDim Preserve compHeight(compCount)
		Set compArray(compCount) = comp

		With comp.BBox
			compWidth(compCount) = .Xmax - .Xmin
			compHeight(compCount) = .Ymax - .Ymin
			currentX = (.Xmin + .Xmax) / 2
			currentY = (.Ymin + .Ymax) / 2
			If currentX < minX Then minX = currentX
			If currentY < minY Then minY = currentY
		End With

		compCount = compCount + 1
	Next comp

	pApp.UnlockServer

	If compCount = 0 Then
		MsgBox "No components selected", vbInformation
		Exit Sub
	End If

	cols = Int(Sqr(compCount))
	If cols < 1 Then cols = 1
	rows = (compCount + cols - 1) / cols

	baseX = minX
	baseY = minY
	rowHeight = 0
	colWidth = 0

	For i = 0 To compCount - 1
		If compHeight(i) > rowHeight Then rowHeight = compHeight(i)
		If compWidth(i) > colWidth Then colWidth = compWidth(i)
	Next i

	pApp.LockServer

	For i = 0 To compCount - 1
		Set comp = compArray(i)
		With comp.BBox
			currentX = (.Xmin + .Xmax) / 2
			currentY = (.Ymin + .Ymax) / 2

			j = i Mod cols
			Dim row As Long
			row = i \ cols

			newX = baseX + (j * (colWidth + padding)) + (compWidth(i) / 2)
			newY = baseY + (row * (rowHeight + padding)) + (compHeight(i) / 2)

			deltaX = newX - currentX
			deltaY = newY - currentY
			comp.Move deltaX, deltaY
		End With
	Next i

	pApp.UnlockServer

	PCBDoc.Refresh
	MsgBox "Arranged " & compCount & " components in grid (" & rows & " rows, " & cols & " cols)", vbInformation

	Exit Sub
ErrorHandler:
	MsgBox "Error: " & Err.Description, vbCritical
End Sub
