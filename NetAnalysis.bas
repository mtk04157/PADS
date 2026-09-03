Sub Main
	On Error GoTo ErrorHandler

	Dim PCBDoc
	Dim net
	Dim netCount As Long
	Dim totalPads As Long
	Dim padCount As Long
	Dim netName As String
	Dim netList As String

	Set PCBDoc = ActiveDocument
	If PCBDoc Is Nothing Then
		MsgBox "Please open PCB file first", vbCritical
		Exit Sub
	End If

	netCount = 0
	totalPads = 0
	netList = ""

	For Each net In PCBDoc.Nets
		netCount = netCount + 1
		padCount = 0

		On Error Resume Next
		padCount = net.PadCount
		On Error GoTo ErrorHandler

		totalPads = totalPads + padCount
		netName = net.Name
		netList = netList & netName & " (" & padCount & " pads)" & vbCrLf
	Next net

	If netCount = 0 Then
		MsgBox "No nets found in design", vbInformation
		Exit Sub
	End If

	MsgBox "Total Networks: " & netCount & vbCrLf & _
		   "Total Pads: " & totalPads & vbCrLf & vbCrLf & _
		   "Networks:" & vbCrLf & netList, vbInformation, "Net Analysis"

	Exit Sub
ErrorHandler:
	MsgBox "Error: " & Err.Description, vbCritical
End Sub
