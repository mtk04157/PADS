Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFFFFFF
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Dim compData() As CompInfo
Dim compCount As Integer
Dim specialCompIndices() As Integer
Dim specialCount As Integer

Type CompInfo
	Name As String
	RefX As Double
	RefY As Double
	MinPadX As Double
	MaxPadX As Double
	MinPadY As Double
	MaxPadY As Double
	PadWidth As Double
	PadHeight As Double
	IsSpecial As Boolean
End Type

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
	specialCount = 0

	Set pApp = PCBDoc.Application
	pApp.LockServer

	Dim comp
	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		ReDim Preserve compData(compCount)
		compData(compCount).Name = comp.Name
		compData(compCount).IsSpecial = (InStr(comp.Name, "U6502") > 0 Or InStr(comp.Name, "U6503") > 0)

		If compData(compCount).IsSpecial Then
			ReDim Preserve specialCompIndices(specialCount)
			specialCompIndices(specialCount) = compCount
			specialCount = specialCount + 1
		End If

		With comp.BBox
			compData(compCount).RefX = (.Xmin + .Xmax) / 2
			compData(compCount).RefY = (.Ymin + .Ymax) / 2
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

	If specialCount > 0 Then
		Call ArrangeWithSpecialHandling(PCBDoc, padding)
	Else
		Call ArrangeStandard(PCBDoc, padding)
	End If

	On Error Resume Next
	Kill ascFile
	Kill tmpMcr
	On Error GoTo 0

	PCBDoc.Refresh
	Dim msg As String
	msg = "Arranged " & compCount & " components with " & Format(padding, "0.0") & "mm pad-to-pad spacing"
	If specialCount > 0 Then
		msg = msg & vbCrLf & "Special: U6502/U6503 outline-to-outline (" & specialCount & " components)"
	End If
	MsgBox msg, vbInformation

	Exit Sub
ErrorHandler:
	MsgBox "Error: " & Err.Description, vbCritical
End Sub

Sub ExportPadData(ByVal PCBDoc, ByVal ascFile As String, ByVal mcrFile As String)
	Dim fh As Integer

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
	Dim fh As Integer
	Dim line As String
	Dim i As Integer
	Dim compName As String
	Dim padX As Double, padY As Double
	Dim padWidth As Double, padHeight As Double

	fh = FreeFile
	On Error Resume Next
	Open ascFile For Input As #fh

	If Err.Number <> 0 Then
		On Error GoTo 0
		Exit Sub
	End If

	Do While Not EOF(fh)
		Line Input #fh, line
		line = Trim(line)

		If Len(line) > 0 And InStr(line, "PAD") > 0 Then
			If InStr(line, "*") > 0 Then
				compName = Trim(Left(line, InStr(line, "*") - 1))

				For i = 0 To compCount - 1
					If compData(i).Name = compName Then
						If InStr(line, "X") > 0 Then
							padX = ExtractCoordinate(line, "X")
							padY = ExtractCoordinate(line, "Y")
							padWidth = ExtractCoordinate(line, "W")
							padHeight = ExtractCoordinate(line, "H")

							If i = 0 Or padX - padWidth / 2 < compData(i).MinPadX Then
								compData(i).MinPadX = padX - padWidth / 2
							End If
							If padX + padWidth / 2 > compData(i).MaxPadX Then
								compData(i).MaxPadX = padX + padWidth / 2
							End If
							If padY - padHeight / 2 < compData(i).MinPadY Then
								compData(i).MinPadY = padY - padHeight / 2
							End If
							If padY + padHeight / 2 > compData(i).MaxPadY Then
								compData(i).MaxPadY = padY + padHeight / 2
							End If

							compData(i).PadWidth = padWidth
							compData(i).PadHeight = padHeight
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
	Dim pos As Integer
	Dim endPos As Integer
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

Sub ArrangeStandard(ByVal PCBDoc, ByVal padding As Double)
	Dim i As Integer
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

	baseX = compData(0).MinPadX
	baseY = compData(0).MinPadY
	nextX = baseX

	For i = 0 To compCount - 1
		Dim compWidth As Double

		compWidth = compData(i).MaxPadX - compData(i).MinPadX
		maxHeight = compData(i).MaxPadY - compData(i).MinPadY

		newX = nextX + compWidth / 2
		newY = baseY + maxHeight / 2

		currentX = compData(i).RefX
		currentY = compData(i).RefY

		deltaX = newX - currentX
		deltaY = newY - currentY

		For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
			If comp.Name = compData(i).Name Then
				comp.Move deltaX, deltaY
				Exit For
			End If
		Next comp

		nextX = nextX + compWidth + padding
	Next i

	pApp.UnlockServer
End Sub

Sub ArrangeWithSpecialHandling(ByVal PCBDoc, ByVal padding As Double)
	Dim i As Integer, j As Integer
	Dim comp
	Dim currentX As Double, currentY As Double
	Dim newX As Double, newY As Double
	Dim deltaX As Double, deltaY As Double
	Dim baseX As Double, baseY As Double
	Dim nextX As Double
	Dim maxHeight As Double
	Dim specialStartIdx As Integer
	Dim pApp

	Set pApp = PCBDoc.Application
	pApp.LockServer

	baseX = compData(0).MinPadX
	baseY = compData(0).MinPadY
	nextX = baseX

	For i = 0 To compCount - 1
		Dim compWidth As Double
		Dim isSpecialGroup As Boolean

		isSpecialGroup = False
		For j = 0 To specialCount - 1
			If specialCompIndices(j) = i Then
				isSpecialGroup = True
				Exit For
			End If
		Next j

		compWidth = compData(i).MaxPadX - compData(i).MinPadX
		maxHeight = compData(i).MaxPadY - compData(i).MinPadY

		newX = nextX + compWidth / 2
		newY = baseY + maxHeight / 2

		currentX = compData(i).RefX
		currentY = compData(i).RefY

		deltaX = newX - currentX
		deltaY = newY - currentY

		For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
			If comp.Name = compData(i).Name Then
				comp.Move deltaX, deltaY
				Exit For
			End If
		Next comp

		If isSpecialGroup Then
			nextX = nextX + compWidth
		Else
			nextX = nextX + compWidth + padding
		End If
	Next i

	pApp.UnlockServer
End Sub
