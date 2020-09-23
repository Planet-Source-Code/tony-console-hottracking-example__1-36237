Attribute VB_Name = "modMain"

'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=
'|| ::Author::          :  TonY Myers           ||
'|| ::Project Name::    :  Console Example      ||
'|| ::Complied          :  25/6/02              ||
'|| ::Created           :  24/6/02 - 01:54am    ||
'|| ::Comments          :  If this code has been||
'|| found on psc.om before i have posted,       ||
'|| sorry :)                                    ||
'||                                             ||
'|| ::NOTES             :    N/A                ||
'||                                             ||
'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=

Option Explicit
Dim i As Integer

Public Sub WriteConsole(logText As String, Optional logColor As Long)
On Error Resume Next

If logColor = 0 Then logColor = &H0&
frmMain.rtfText.SelStart = Len(frmMain.rtfText)
frmMain.rtfText.SelColor = logColor
frmMain.rtfText.SelText = logText & vbCrLf
frmMain.rtfText.SelStart = Len(frmMain.rtfText)
DoEvents
End Sub
