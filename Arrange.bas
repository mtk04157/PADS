Sub Main
	On Error GoTo ErrorHandler

	Dim PCBDoc
	Dim comp
	Dim i As Long
	Dim spacing As Double
	Dim pApp
	Dim compName As String
	Dim successCount As Long
	Dim skipList As String

	Set PCBDoc = ActiveDocument
	If PCBDoc Is Nothing Then
		MsgBox "Please open PCB file first", vbCritical
		Exit Sub
	End If

	spacing = 2.0
	i = 0
	successCount = 0
	skipList = ""

	Set pApp = PCBDoc.Application
	pApp.LockServer

	For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
		compName = comp.Name
		On Error Resume Next
		comp.Move i * spacing, 0
		If Err.Number <> 0 Then
			skipList = skipList & compName & vbCrLf
			Err.Clear
		Else
			successCount = successCount + 1
		End If
		On Error GoTo ErrorHandler
		i = i + 1
	Next comp

	pApp.UnlockServer

	PCBDoc.Refresh
	
	Dim msg As String
	msg = "Arranged " & successCount & " components"
	If Len(skipList) > 0 Then
		msg = msg & vbCrLf & "Skipped: " & vbCrLf & skipList
	End If
	MsgBox msg, vbInformation

	Exit Sub
ErrorHandler:
	pApp.UnlockServer
	MsgBox "Error: " & Err.Description, vbCritical
End Sub
