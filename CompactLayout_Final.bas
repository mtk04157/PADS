Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFFFFFF
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private compName() As String
Private compRefX() As Double
Private compRefY() As Double
Private compMinPadX() As Double
Private compMaxPadX() As Double
Private compMinPadY() As Double
Private compMaxPadY() As Double
Private compCount As Long

Sub Main
	On Error GoTo ErrorHandler

	Dim PCBDoc
	Dim comp
	Dim i As Long
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

	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		ReDim Preserve compName(compCount)
		ReDim Preserve compRefX(compCount)
		ReDim Preserve compRefY(compCount)
		ReDim Preserve compMinPadX(compCount)
		ReDim Preserve compMaxPadX(compCount)
		ReDim Preserve compMinPadY(compCount)
		ReDim Preserve compMaxPadY(compCount)

		compName(compCount) = comp.Name
		With comp.BBox
			compRefX(compCount) = (.Xmin + .Xmax) / 2
			compRefY(compCount) = (.Ymin + .Ymax) / 2
		End With

		compMinPadX(compCount) = 999999
		compMaxPadX(compCount) = -999999
		compMinPadY(compCount) = 999999
		compMaxPadY(compCount) = -999999

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
	Call ArrangeComponents(PCBDoc, padding)

	On Error Resume Next
	Kill ascFile
	Kill tmpMcr
	On Error GoTo 0

	PCBDoc.Refresh
	MsgBox "Arranged " & compCount & " components", vbInformation

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
	On Error GoTo 0

	Do While Not EOF(fh)
		Line Input #fh, line
		line = Trim(line)

		If Len(line) > 0 And InStr(line, "PAD") > 0 Then
			If InStr(line, "*") > 0 Then
				compNameStr = Trim(Left(line, InStr(line, "*") - 1))

				For i = 0 To compCount - 1
					If compName(i) = compNameStr Then
						padX = ExtractValue(line, "X")
						padY = ExtractValue(line, "Y")
						padWidth = ExtractValue(line, "W")
						padHeight = ExtractValue(line, "H")

						If padX - padWidth / 2 < compMinPadX(i) Then
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
						Exit For
					End If
				Next i
			End If
		End If
	Loop

	Close #fh
End Sub

Function ExtractValue(ByVal line As String, ByVal key As String) As Double
	Dim pos As Long
	Dim endPos As Long
	Dim valueStr As String

	pos = InStr(line, key & " ")
	If pos = 0 Then pos = InStr(line, key & "=")
	If pos = 0 Then pos = InStr(line, key)

	If pos > 0 Then
		valueStr = Mid(line, pos + 1)
		valueStr = Trim(valueStr)
		endPos = InStr(valueStr, " ")
		If endPos > 0 Then
			valueStr = Left(valueStr, endPos - 1)
		End If
		On Error Resume Next
		ExtractValue = CDbl(valueStr)
		On Error GoTo 0
	End If
End Function

Sub ArrangeComponents(ByVal PCBDoc, ByVal padding As Double)
	Dim i As Long
	Dim comp
	Dim currentX As Double, currentY As Double
	Dim newX As Double, newY As Double
	Dim deltaX As Double, deltaY As Double
	Dim baseX As Double, baseY As Double
	Dim nextX As Double
	Dim compWidth As Double
	Dim pApp

	Set pApp = PCBDoc.Application
	pApp.LockServer

	baseX = compMinPadX(0)
	baseY = compMinPadY(0)
	nextX = baseX

	For i = 0 To compCount - 1
		compWidth = compMaxPadX(i) - compMinPadX(i)

		newX = nextX + compWidth / 2
		newY = baseY

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
