VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alfred C. Toepfer"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   5535
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   5535
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Encriptar"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Desencriptar"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   3120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ConnectionStringEncryptionKey As String = "“∂st~'KÇøåØ^Å˘"



Private Sub cmdLoad_Click()
    Dim sText As String
    
    sText = txtInput.Text
    sText = RC4(Decode(sText), ConnectionStringEncryptionKey)
    txtOutput.Text = sText
    
End Sub

Private Sub cmdSave_Click()
    Dim sText As String
    
    sText = txtInput.Text
    sText = Encode(RC4(sText, ConnectionStringEncryptionKey))
    txtOutput.Text = sText
    
End Sub

Public Function Decode(Value As String) As String
    Dim Loop1 As Long
    
    For Loop1 = 1 To Len(Value) Step 2
        Decode = Decode & Chr(Val("&H" & Mid(Value, Loop1, 2)))
    Next
End Function
Public Function Encode(Value As String) As String
    Dim Loop1 As Long
    Dim SingleValue As String
    
    For Loop1 = 1 To Len(Value)
        SingleValue = Hex(Asc(Mid(Value, Loop1, 1)))
        If Len(SingleValue) = 1 Then SingleValue = "0" & SingleValue
        Encode = Encode & SingleValue
    Next
End Function

