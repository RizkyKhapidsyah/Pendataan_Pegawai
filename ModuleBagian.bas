Attribute VB_Name = "Module1"
Option Explicit

Public CN As New ADODB.Connection
Public RS As New ADODB.Recordset
Global X As String
Public Pesan As Integer
Global Objek As Control

Public Sub Nyambungg()
If CN.State = adStateOpen Then CN.Close
    CN.CursorLocation = adUseClient
    CN.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=DBPendataanPegawai"
End Sub

Public Sub PusatError()
MsgBox "Maaf, terjadi kesalahan!" & vbCrLf & vbCrLf & _
        "Error : " & Err.Description & vbCrLf & _
        "Code : " & Err.Number, vbCritical + vbOKOnly, "Error"
End Sub
