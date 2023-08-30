VERSION 5.00
Begin VB.Form Inicio 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "To-Do List"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtNome 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdCommand 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insira a senha"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Label lblNameAdvise 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insira o seu nome"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   1635
   End
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NameInput As String
Dim PassInput As String


Private Sub txtNome_Click()
    txtNome.Text = ""
End Sub

Private Sub txtNome_Change()
    NameInput = txtNome.Text
End Sub

Private Sub cmdCommand_Click()
    Dim Style, Response
    Style = vbExclamation
        If NameInput = "Teste" And PassInput = "aaa123" Then
            MsgBox ("Login Realizado com sucesso!!!")
            MenuForm.Show
            Me.Hide
        ElseIf NameInput <> "Teste" Then
            Response = MsgBox("Usuário Inexistente, tente novamente!!!", Style)
        ElseIf PassInput <> "aaa123" Then
            Response = MsgBox("Senha errada, tente novamente!!!", Style)
        End If
End Sub

Private Sub txtNome_LostFocus()
    If txtNome.Text = "" Then
        txtNome.BackColor = RGB(255, 0, 0)
    Else
        txtNome.BackColor = RGB(255, 255, 0)
    End If
End Sub

Private Sub txtPassword_Click()
    txtPassword.Text = ""
End Sub

Private Sub txtPassword_Change()
    txtPassword.PasswordChar = "*"
    PassInput = txtPassword.Text
    
End Sub

Private Sub txtPassword_LostFocus()
    If txtPassword.Text = "" Then
       txtPassword.BackColor = RGB(255, 0, 0)
    Else
        txtPassword.BackColor = RGB(255, 255, 0)
    End If
End Sub
