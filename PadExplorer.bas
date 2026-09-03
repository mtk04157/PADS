Sub Main
	On Error GoTo ErrorHandler

	Dim PCBDoc
	Dim comp
	Dim pad
	Dim padCount As Integer
	Dim compName As String
	Dim padInfo As String
	Dim pApp

	Set PCBDoc = ActiveDocument
	If PCBDoc Is Nothing Then
		MsgBox "Please open PCB file first", vbCritical
		Exit Sub
	End If

	pApp = PCBDoc.Application
	pApp.LockServer

	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		compName = comp.Name
		padCount = 0
		padInfo = "Component: " & compName & vbCrLf

		On Error Resume Next

		If Not comp.Pads Is Nothing Then
			padInfo = padInfo & "Pads collection found:" & vbCrLf
			For Each pad In comp.Pads
				padCount = padCount + 1
				padInfo = padInfo & "  Pad " & padCount & ": " & pad.Name & vbCrLf
			Next pad
		End If

		On Error GoTo 0

		padInfo = padInfo & "Total Pads: " & padCount & vbCrLf

		MsgBox padInfo, vbInformation, "Pad Info"

		pApp.UnlockServer
		Exit Sub

	Next comp

	MsgBox "No components selected", vbInformation

	pApp.UnlockServer

	Exit Sub
ErrorHandler:
	On Error GoTo 0
	MsgBox "Error at pad exploration: " & Err.Description & vbCrLf & _
		   "Pad access may require different approach (ASCII export method)", vbCritical
End Sub
