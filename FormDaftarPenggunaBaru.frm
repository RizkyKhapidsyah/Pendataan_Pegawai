VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormDaftarPenggunaBaru 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Pengguna Baru"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Sylfaen"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormDaftarPenggunaBaru.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   1800
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin VB.CommandButton cmBatal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Batal"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmOK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&OK"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox textKonfirmasiPassword 
      Height          =   390
      Left            =   2520
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox textPasswordBaru 
      Height          =   390
      Left            =   2520
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox textNamaPenggunaBaru 
      Height          =   390
      Left            =   2520
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   2400
      TabIndex        =   7
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Konfirmasi Password"
      Height          =   270
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   2400
      TabIndex        =   4
      Top             =   600
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password Baru"
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pengguna Baru"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1770
   End
End
Attribute VB_Name = "FormDaftarPenggunaBaru"
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
    textPasswordBaru.PasswordChar = "*"
    textKonfirmasiPassword.PasswordChar = "*"
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

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmOK_Click()
If textNamaPenggunaBaru.Text = "" Then
    MsgBox "Silahkan isi nama pengguna yang ingin didaftarkan!", vbExclamation + vbOKOnly, ""
    textNamaPenggunaBaru.SetFocus
ElseIf textPasswordBaru.Text = "" Then
    MsgBox "Silahkan isi Password baru Anda!", vbExclamation + vbOKOnly, ""
    textPasswordBaru.SetFocus
ElseIf textKonfirmasiPassword.Text = "" Then
    MsgBox "Silahkan konfirmasikan password baru Anda!", vbExclamation + vbOKOnly, ""
    textKonfirmasiPassword.SetFocus
ElseIf textPasswordBaru.Text <> textKonfirmasiPassword.Text Then
    MsgBox "Maaf, password baru tidak sesuai dengan konfirmasi password Anda!", vbCritical + vbOKOnly, "Error"
    textPasswordBaru.SetFocus
ElseIf textPasswordBaru.MaxLength <= 5 Then
    MsgBox "Password setidaknya minimal 6 karakter", vbExclamation + vbOKOnly, ""
    textPasswordBaru.SetFocus
Else
    Pesan = MsgBox("Data sudah benar. Apakah Anda yakin ingin mendaftarkan pengguna ini?", vbQuestion + vbYesNo, "Konfirmasi")
    If Pesan = vbYes Then
        With Adodc1
            .Recordset.AddNew
            .Recordset.Fields(0).Value = textNamaPenggunaBaru.Text
            .Recordset.Fields(1).Value = textPasswordBaru.Text
            .Recordset.Update
            .Refresh
        End With
            MsgBox "Pengguna dengan Nama : '" & textNamaPenggunaBaru.Text & "' telah didaftarkan!", vbInformation + vbOKOnly, "Berhasil"
            KosongkanInput
            Unload Me
            FormLogin.AturKontrol
    End If
End If
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
