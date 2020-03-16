Option Explicit

Dim WshShell, oExec, BtnCode
Dim Failed, Line, VolumeTypes, Healthy, Message

Set Healthy = New RegExp
Set VolumeTypes = New RegExp

Healthy.Pattern = "Healthy"
VolumeTypes.Pattern = "Mirror|RAID-5|Partition"

Message = ""
Failed = 0

Set WshShell = WScript.CreateObject("WScript.Shell")
Set oExec = WshShell.Exec("%comspec% /C echo list volume | %WINDIR%\SYSTEM32\DISKPART.EXE")

If oExec.StdOut.AtEndOfStream Then
	Failed = 1
	Message = "No output from DISKPART"
Else
	While Not oExec.StdOut.AtEndOfStream
		Line = oExec.StdOut.ReadLine

		If VolumeTypes.Test(Line) Then

			If Healthy.Test(Line) Then
				Message = Message & "Healthy:"
				Message = Message & Line & vbCrLf
			Else
				Failed = 1
				Message = Message & "Faulty:"
				Message = Message & Line & vbCrLf
			End If
		End If
	WEnd
End If

if Failed = 1 Then
	BtnCode = WshShell.Popup("Raid or Volume Failure!" & vbCrLf & Message)
Else
	BtnCode = WshShell.Popup(Message)
End If

