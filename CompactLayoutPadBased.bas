Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFFFFFF
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Dim compList() As String
Dim compPads() As String
Dim compMinX() As Double
Dim compMinY() As Double
Dim compMaxX() As Double
Dim compMaxY() As Double

Sub Main
	On Error GoTo ErrorHandler

	Dim PCBDoc
	Dim comp
	Dim compCount As Long
	Dim ascFile As String
	Dim padding As Double

	Set PCBDoc = ActiveDocument
	If PCBDoc Is Nothing Then
		MsgBox "Please open PCB file first", vbCritical
		Exit Sub
	End If

	compCount = 0
	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		ReDim Preserve compList(compCount)
		compList(compCount) = comp.Name
		compCount = compCount + 1
	Next comp

	If compCount = 0 Then
		MsgBox "No components selected", vbInformation
		Exit Sub
	End If

	ascFile = PCBDoc.path & "\" & PCBDoc.Name & "_pad_info.asc"
	PCBDoc.ExportASCII(ascFile, ppcbASCIISectionPads Or ppcbASCIISectionConnections)

	Call ParsePadData(ascFile, compCount)

	Call CalculateCompactLayout(PCBDoc, compCount, 0.2)

	Kill ascFile

	PCBDoc.Refresh
	MsgBox "Compact layout complete for " & compCount & " components with 0.2mm pad-to-pad spacing", vbInformation

	Exit Sub
ErrorHandler:
	MsgBox "Error: " & Err.Description, vbCritical
End Sub

Sub ParsePadData(ByVal ascFile As String, ByVal compCount As Long)
	Dim fh As Long
	Dim line As String
	Dim i As Long
	Dim padX As Double, padY As Double
	Dim compName As String
	Dim minX As Double, minY As Double, maxX As Double, maxY As Double

	ReDim compMinX(compCount - 1)
	ReDim compMinY(compCount - 1)
	ReDim compMaxX(compCount - 1)
	ReDim compMaxY(compCount - 1)

	For i = 0 To compCount - 1
		compMinX(i) = 999999
		compMinY(i) = 999999
		compMaxX(i) = -999999
		compMaxY(i) = -999999
	Next i

	fh = FreeFile
	On Error Resume Next
	Open ascFile For Input As #fh

	Do While Not EOF(fh)
		Line Input #fh, line
		If Len(Trim(line)) > 0 Then
			If InStr(line, "/") > 0 And InStr(line, " ") > 0 Then
				compName = Left(line, InStr(line, "/") - 1)
				compName = Trim(compName)

				For i = 0 To compCount - 1
					If compList(i) = compName Then
						If InStr(line, " X ") > 0 Then
							padX = CDbl(Mid(line, InStr(line, " X ") + 3, 10))
							padY = CDbl(Mid(line, InStr(line, " Y ") + 3, 10))

							If padX < compMinX(i) Then compMinX(i) = padX
							If padY < compMinY(i) Then compMinY(i) = padY
							If padX > compMaxX(i) Then compMaxX(i) = padX
							If padY > compMaxY(i) Then compMaxY(i) = padY
						End If
						Exit For
					End If
				Next i
			End If
		End If
	Loop

	Close #fh
	On Error GoTo 0
End Sub

Sub CalculateCompactLayout(ByVal PCBDoc, ByVal compCount As Long, ByVal padding As Double)
	Dim i As Long, j As Long
	Dim comp
	Dim currentX As Double, currentY As Double
	Dim newX As Double, newY As Double
	Dim deltaX As Double, deltaY As Double
	Dim baseX As Double, baseY As Double
	Dim compWidth As Double, compHeight As Double
	Dim pApp

	Set pApp = PCBDoc.Application
	pApp.LockServer

	baseX = compMinX(0)
	baseY = compMinY(0)

	j = 0
	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		For i = 0 To compCount - 1
			If compList(i) = comp.Name Then
				With comp.BBox
					currentX = (.Xmin + .Xmax) / 2
					currentY = (.Ymin + .Ymax) / 2
				End With

				compWidth = compMaxX(i) - compMinX(i)
				compHeight = compMaxY(i) - compMinY(i)

				newX = baseX + (j * (compWidth + padding)) + (compWidth / 2)
				newY = baseY + (compHeight / 2)

				deltaX = newX - currentX
				deltaY = newY - currentY
				comp.Move deltaX, deltaY

				j = j + 1
				Exit For
			End If
		Next i
	Next comp

	pApp.UnlockServer
End Sub
