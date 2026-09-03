Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFFFFFF
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Dim compName() As String
Dim compRefX() As Double
Dim compRefY() As Double
Dim compMinPadX() As Double
Dim compMaxPadX() As Double
Dim compMinPadY() As Double
Dim compMaxPadY() As Double
Dim compPadWidth() As Double
Dim compPadHeight() As Double
Dim compCount As Long

Sub Main
	On Error GoTo ErrorHandler

	Dim PCBDoc
	Dim ascFile As String
	Dim tmpMcr As String
	Dim padding As Double
	Dim pApp

	Set PCBDoc = ActiveDocument
	If PCBDoc Is Nothing Then
		MsgBox "Please open PCB file first", vbCritical
		Exit Sub
	End If

	padding = 0.2
	compCount = 0

	Set pApp = PCBDoc.Application
	pApp.LockServer

	Dim comp
	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		ReDim Preserve compName(compCount)
		ReDim Preserve compRefX(compCount)
		ReDim Preserve compRefY(compCount)
		ReDim Preserve compMinPadX(compCount)
		ReDim Preserve compMaxPadX(compCount)
		ReDim Preserve compMinPadY(compCount)
		ReDim Preserve compMaxPadY(compCount)
		ReDim Preserve compPadWidth(compCount)
		ReDim Preserve compPadHeight(compCount)

		compName(compCount) = comp.Name

		With comp.BBox
			compRefX(compCount) = (.Xmin + .Xmax) / 2
			compRefY(compCount) = (.Ymin + .Ymax) / 2
		End With

		compCount = compCount + 1
	Next comp

	pApp.UnlockServer

	If compCount = 0 Then
		MsgBox "No components selected", vbInformation
		Exit Sub
	End If

	ascFile = Environ$("TEMP") & "\_tmp_pads_exact.asc"
	tmpMcr = Environ$("TEMP") & "\_tmp_export_pads.mcr"

	Call ExportPadData(PCBDoc, ascFile, tmpMcr)

	Call ParsePadCoordinates(ascFile)

	Call ArrangeWithPadSpacing(PCBDoc, padding)

	On Error Resume Next
	Kill ascFile
	Kill tmpMcr
	On Error GoTo 0

	PCBDoc.Refresh
	MsgBox "Arranged " & compCount & " components with " & Format(padding, "0.0") & "mm pad-to-pad spacing", vbInformation

	Exit Sub
ErrorHandler:
	MsgBox "Error: " & Err.Description, vbCritical
End Sub

Sub ExportPadData(ByVal PCBDoc, ByVal ascFile As String, ByVal mcrFile As String)
	Dim fh As Long

	fh = FreeFile
	Open mcrFile For Output As #fh
	Print #fh, "Application.ExportDocument(""" & ascFile & """)"
	Print #fh, "ASCIIOutDlg.Selections.Checked(3) = true"
	Print #fh, "ASCIIOutDlg.Selections.Checked(6) = true"
	Print #fh, "ASCIIOutDlg.Units = ""Current"""
	Print #fh, "ASCIIOutDlg.Ok.Click()"
	Close #fh

	RunMacro(mcrFile, "")
End Sub

Sub ParsePadCoordinates(ByVal ascFile As String)
	Dim fh As Long
	Dim line As String
	Dim i As Long
	Dim compNameStr As String
	Dim padX As Double, padY As Double
	Dim padWidth As Double, padHeight As Double

	fh = FreeFile
	On Error Resume Next
	Open ascFile For Input As #fh

	If Err.Number <> 0 Then
		On Error GoTo 0
		MsgBox "Cannot open ASCII export file", vbCritical
		Exit Sub
	End If

	Do While Not EOF(fh)
		Line Input #fh, line
		line = Trim(line)

		If Len(line) > 0 And InStr(line, "PAD") > 0 Then
			If InStr(line, "*") > 0 Then
				compNameStr = Trim(Left(line, InStr(line, "*") - 1))

				For i = 0 To compCount - 1
					If compName(i) = compNameStr Then
						If InStr(line, "X") > 0 Then
							padX = ExtractCoordinate(line, "X")
							padY = ExtractCoordinate(line, "Y")
							padWidth = ExtractCoordinate(line, "W")
							padHeight = ExtractCoordinate(line, "H")

							If i = 0 Or padX - padWidth / 2 < compMinPadX(i) Then
								compMinPadX(i) = padX - padWidth / 2
							End If
							If padX + padWidth / 2 > compMaxPadX(i) Then
								compMaxPadX(i) = padX + padWidth / 2
							End If
							If padY - padHeight / 2 < compMinPadY(i) Then
								compMinPadY(i) = padY - padHeight / 2
							End If
							If padY + padHeight / 2 > compMaxPadY(i) Then
								compMaxPadY(i) = padY + padHeight / 2
							End If

							compPadWidth(i) = padWidth
							compPadHeight(i) = padHeight
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

Function ExtractCoordinate(ByVal line As String, ByVal coordChar As String) As Double
	Dim pos As Long
	Dim endPos As Long
	Dim valueStr As String

	pos = InStr(line, coordChar & " ")
	If pos = 0 Then pos = InStr(line, coordChar & "=")
	If pos = 0 Then pos = InStr(line, coordChar)

	If pos > 0 Then
		valueStr = Mid(line, pos + 1)
		valueStr = Trim(valueStr)
		endPos = InStr(valueStr, " ")
		If endPos > 0 Then
			valueStr = Left(valueStr, endPos - 1)
		End If
		On Error Resume Next
		ExtractCoordinate = CDbl(valueStr)
		On Error GoTo 0
	End If
End Function

Sub ArrangeWithPadSpacing(ByVal PCBDoc, ByVal padding As Double)
	Dim i As Long
	Dim comp
	Dim currentX As Double, currentY As Double
	Dim newX As Double, newY As Double
	Dim deltaX As Double, deltaY As Double
	Dim baseX As Double, baseY As Double
	Dim nextX As Double
	Dim maxHeight As Double
	Dim pApp

	Set pApp = PCBDoc.Application
	pApp.LockServer

	baseX = compMinPadX(0)
	baseY = compMinPadY(0)
	nextX = baseX

	For i = 0 To compCount - 1
		Dim compWidth As Double

		compWidth = compMaxPadX(i) - compMinPadX(i)
		maxHeight = compMaxPadY(i) - compMinPadY(i)

		newX = nextX + compWidth / 2
		newY = baseY + maxHeight / 2

		currentX = compRefX(i)
		currentY = compRefY(i)

		deltaX = newX - currentX
		deltaY = newY - currentY

		For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
			If comp.Name = compName(i) Then
				comp.Move deltaX, deltaY
				Exit For
			End If
		Next comp

		nextX = nextX + compWidth + padding
	Next i

	pApp.UnlockServer
End Sub
