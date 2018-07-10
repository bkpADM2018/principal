Attribute VB_Name = "Main"

Public sFile As String
Public Function SaveConfigFile() As Boolean
      Dim fso As Object
      Dim f As Object
      Dim sText As String

10    On Error GoTo ErrHandler

20    Set fso = CreateObject("Scripting.FileSystemObject")
30    Set f = fso.CreateTextFile(sFile, True)

40    sText = frmMain.txtInput.Text
50    sText = Encode(RC4(sText, ConnectionStringEncryptionKey))
60    f.Write sText
70    f.Close
80    SaveConfigFile = True

90    Set fso = Nothing
100   Set f = Nothing

110   GoTo Exit_Function

ErrHandler:
120   GrabarLog "Error Number: " & Err.Number & vbCrLf & _
                "Error Description: " & Err.Description & vbCrLf & _
                "Line Number: " & Erl & vbCrLf & _
                "Module Name: SaveConfigFile de Módulo modGral"
130   SaveConfigFile = False
Exit_Function:
End Function

Public Function LoadConfigFile() As Boolean

Dim fso As Object
Dim f As Object
Dim sText As String

On Error GoTo ErrHandler
LoadConfigFile = True

Set fso = CreateObject("Scripting.FileSystemobject")
If Not fso.FileExists(sFile) Then
    SaveConfigFile
End If

Set f = fso.OpenTextFile(sFile, 1)
If Not f.atendofstream Then sText = RC4(Decode(f.ReadAll), ConnectionStringEncryptionKey)
f.Close
frmMain.txtOutput.Text = sText

Set f = Nothing
Set fso = Nothing

Set oXML = Nothing

GoTo Exit_Function

ErrHandler:
GrabarLog "Error Number: " & Err.Number & vbCrLf & _
          "Error Description: " & Err.Description & vbCrLf & _
          "Line Number: " & Erl & vbCrLf & _
          "Module Name: LoadConfigFile de Módulo modGral"
LoadConfigFile = False
Exit_Function:
End Function

Public Sub GrabarLog(ByVal sValue As String)
Dim ff As Integer

ff = FreeFile()
If sLogPath = "" Then
    sLogPath = App.Path & "\ToepferPass.log"
End If
If Dir(sLogPath) = "" Then
    Open sLogPath For Output As ff
Else
    Open sLogPath For Append As ff
End If

Print #ff, Now() & " - " & App.ProductName & " - " & sValue

Close ff
MsgBox "Se ha producido un error. Por favor consulte el log."
End Sub

