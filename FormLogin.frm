VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Sylfaen"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   2760
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmDaftar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Daftar"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmKeluar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmLogin 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox textPassword 
      Height          =   390
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox textPengguna 
      Height          =   390
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pengguna"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    Nyambungg
    With Adodc1
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From tbLogin"
        .Refresh
    End With
    KosongkanInput
    textPassword.PasswordChar = "*"
End Sub
Sub KosongkanInput()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .Text = ""
                .MaxLength = 254
            End With
        End If
    Next
End Sub

Private Sub cmDaftar_Click()
    FormDaftarPenggunaBaru.Show vbModal, Me
End Sub

Private Sub cmKeluar_Click()
Pesan = MsgBox("Anda yakin ingin keluar?", vbQuestion + vbYesNo, "Keluar?")
        If Pesan = vbYes Then End
End Sub

Private Sub cmLogin_Click()
    X = "Select * From tbLogin where Pengguna = '" & textPengguna.Text & "' and Password = '" & textPassword.Text & "'"
        Set RS = CN.Execute(X)
        If textPengguna.Text = "" Then
            MsgBox " Silahkan isi nama pengguna Anda!", vbExclamation + vbOKOnly, "Nama Pengguna?"
            textPengguna.SetFocus
        ElseIf textPassword.Text = "" Then
            MsgBox "Silahkan isi password Anda!", vbExclamation + vbOKOnly, "Password"
            textPassword.SetFocus
        Else
            If Not RS.EOF Then
                Me.Hide
                With FormUtama
                    .Show
                    .SetFocus
                End With
            Else
                MsgBox "Mohon perika kembali Nama Pengguna dan Password Anda!", vbExclamation + vbOKOnly, "Salah!"
                KosongkanInput
                textPengguna.SetFocus
            End If
        End If
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

Private Sub Form_Unload(Cancel As Integer)
Pesan = MsgBox("Anda yakin ingin keluar?", vbQuestion + vbYesNo, "Keluar?")
        If Pesan = vbYes Then
            End
        Else
            Cancel = 1
        End If
End Sub
