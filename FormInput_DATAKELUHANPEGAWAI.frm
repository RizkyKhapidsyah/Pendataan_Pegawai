VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormInput_DATAKELUHANPEGAWAI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Data Keluhan"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Sylfaen"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormInput_DATAKELUHANPEGAWAI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   5055
      Left            =   600
      TabIndex        =   6
      Top             =   0
      Width           =   4455
      Begin VB.ComboBox cmbNIP 
         Height          =   390
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox TextNama 
         Height          =   390
         Left            =   1440
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox textBagian 
         Height          =   390
         Left            =   1440
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox textJabatan 
         Height          =   390
         Left            =   1440
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox textTanggal 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   1440
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox textBulan 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   2160
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox textTahun 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   2880
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox textKeluhan 
         Height          =   855
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "FormInput_DATAKELUHANPEGAWAI.frx":000C
         Top             =   3120
         Width           =   2895
      End
      Begin VB.CommandButton cmDeteksiTanggal 
         Caption         =   "<>"
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox textJamPermisi 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   1440
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmDeteksiJamPermisi 
         Caption         =   "<>"
         Height          =   375
         Left            =   2880
         TabIndex        =   10
         Top             =   4080
         Width           =   375
      End
      Begin VB.TextBox textLamaPermisi 
         Height          =   390
         Left            =   1440
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   4560
         Width           =   1335
      End
      Begin VB.ComboBox cmbSatuanLamaPermisi 
         Height          =   390
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   4560
         Width           =   1455
      End
      Begin VB.ComboBox cmbJenisPegawai 
         Height          =   390
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIP"
         Height          =   270
         Left            =   870
         TabIndex        =   38
         Top             =   720
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   37
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   270
         Left            =   720
         TabIndex        =   36
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   35
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bagian"
         Height          =   270
         Left            =   675
         TabIndex        =   34
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   33
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         Height          =   270
         Left            =   615
         TabIndex        =   32
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   31
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   270
         Left            =   585
         TabIndex        =   30
         Top             =   2640
         Width           =   630
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   29
         Top             =   2640
         Width           =   45
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keluhan"
         Height          =   270
         Left            =   540
         TabIndex        =   28
         Top             =   3120
         Width           =   675
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   27
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Permisi"
         Height          =   270
         Left            =   255
         TabIndex        =   26
         Top             =   4080
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   25
         Top             =   4080
         Width           =   45
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lama Permisi"
         Height          =   270
         Left            =   120
         TabIndex        =   24
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   23
         Top             =   4560
         Width           =   45
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   22
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Pegawai"
         Height          =   270
         Left            =   135
         TabIndex        =   21
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmBaru 
      Caption         =   "&Baru"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmManage 
      Caption         =   "&Manage"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmTutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox textKode 
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   8400
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc AdodcNIP 
      Height          =   330
      Left            =   120
      Top             =   7920
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc AdodcUtama 
      Height          =   330
      Left            =   120
      Top             =   7560
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
   Begin VB.Menu MenuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu menuVIK 
         Caption         =   "Verifikasi Input Kosong"
      End
   End
End
Attribute VB_Name = "FormInput_DATAKELUHANPEGAWAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
Nyambungg
With AdodcUtama
    .ConnectionString = CN.ConnectionString
    .RecordSource = "Select * From TbKeluhanPegawai"
    .Refresh
End With
With AdodcNIP
    .ConnectionString = CN.ConnectionString
    .RecordSource = "Select * From TbDataIdentitasPegawai"
    .Refresh
End With
For Each Objek In Me
    If TypeName(Objek) = "TextBox" Then
        With Objek
            .Text = ""
            .MaxLength = 254
        End With
    End If
Next
    With cmbJenisPegawai
        .Clear
        .AddItem "Tetap", 0
        .AddItem "Honor", 1
        .AddItem "Kontrak", 2
        .ListIndex = 0
    End With
    With cmbNIP
        .Clear
        .AddItem "-"
        Do Until AdodcNIP.Recordset.EOF
            .AddItem AdodcNIP.Recordset.Fields(0).Value
            AdodcNIP.Recordset.MoveNext
        Loop
    End With
    TextNama.Locked = True
    textBagian.Locked = True
    textJabatan.Locked = True
    textTanggal.Text = Day(Date)
    textBulan.Text = Month(Date)
    textTahun.Text = Year(Date)
    textKode.Text = Second(Time) & "KP" & Hour(Time) & "R-PLN" & Minute(Time)
    With cmbSatuanLamaPermisi
        .Clear
        .AddItem "Menit", 0
        .AddItem "Jam", 1
        .AddItem "Hari", 2
        .AddItem "Minggu", 3
        .ListIndex = 2
    End With
End Sub
Sub KosongkanInput()
For Each Objek In Me
    If TypeName(Objek) = "TextBox" Then Objek.Text = ""
Next
End Sub
Sub AktifkanInput()
For Each Objek In Me
    If TypeName(Objek) = "TextBox" Then
        With Objek
            .Enabled = True
            .BackColor = vbWhite
        End With
    ElseIf TypeName(Objek) = "ComboBox" Then
        With Objek
            .Enabled = True
            .BackColor = vbWhite
        End With
    End If
Next
    menuVIK.Enabled = True
End Sub
Sub NonAktifkanInput()
For Each Objek In Me
    If TypeName(Objek) = "TextBox" Then
        With Objek
            .Enabled = False
            .BackColor = Me.BackColor
        End With
    ElseIf TypeName(Objek) = "ComboBox" Then
        With Objek
            .Enabled = False
            .BackColor = Me.BackColor
        End With
    End If
Next
    menuVIK.Enabled = False
End Sub


Private Sub cmBaru_Click()
    cmBaru.Enabled = False
    cmSimpan.Enabled = True
    cmReset.Enabled = True
    AktifkanInput
    textTanggal.Text = Day(Date)
    textBulan.Text = Month(Date)
    textTahun.Text = Year(Date)
    textJamPermisi.Text = Time
    cmbJenisPegawai.SetFocus
End Sub

Private Sub cmbJenisPegawai_Click()
    If cmbJenisPegawai.ListIndex = 1 Then
        With cmbNIP
            .Text = "-"
            .Enabled = False
        End With
    Else
        With cmbNIP
            .Enabled = True
        End With
    End If
End Sub

Private Sub cmbNIP_Click()
    AdodcNIP.Recordset.MoveFirst
    If RS.State = 1 Then RS.Close
    RS.Open "select * from tbdataidentitaspegawai where NIP = '" & cmbNIP.Text & "'", CN, 3, 3
    If Not RS.EOF Then
        TextNama.Text = RS.Fields("Nama_Pegawai")
        textBagian.Text = RS.Fields("Bagian")
        textJabatan.Text = RS.Fields("Jabatan")
    End If
End Sub

Private Sub cmDeteksiJamPermisi_Click()
    textJamPermisi.Text = Time
End Sub

Private Sub cmDeteksiTanggal_Click()
    textTanggal.Text = Day(Date)
    textBulan.Text = Month(Date)
    textTahun.Text = Year(Date)
End Sub

Private Sub cmDeteksiTanggal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub cmReset_Click()
    KosongkanInput
    cmbNIP.SetFocus
    textKode.Text = Second(Time) & "KP" & Hour(Time) & "R-PLN" & Minute(Time)
    textTanggal.Text = Day(Date)
    textBulan.Text = Month(Date)
    textTahun.Text = Year(Date)
    textJamPermisi.Text = Time
End Sub

Private Sub cmSimpan_Click()
On Error GoTo HancurkanError
If TextNama.Text = "" Or textBagian.Text = "" Or textJabatan.Text = "" Then
    MsgBox "Silahkan pilih NIP pegawai!", vbExclamation + vbOKOnly, ""
    cmbNIP.SetFocus
ElseIf textTanggal.Text = "" Then
    MsgBox "Silahkan isi tanggal!", vbExclamation + vbOKOnly, ""
    textTanggal.SetFocus
ElseIf textBulan.Text = "" Then
    MsgBox "Silahkan isi Bulan !", vbExclamation + vbOKOnly, ""
    textBulan.SetFocus
ElseIf textTahun.Text = "" Then
    MsgBox "Silahkan isi tahun!", vbExclamation + vbOKOnly, ""
    textTahun.SetFocus
ElseIf textKeluhan.Text = "" Then
    MsgBox "Silahkan isi keluhan pegawai!", vbExclamation + vbOKOnly, ""
    textKeluhan.SetFocus
ElseIf textJamPermisi.Text = "" Then
    MsgBox "Silahkan isi jam permisi pegawai!", vbExclamation + vbOKOnly, ""
    textJamPermisi.SetFocus
ElseIf textLamaPermisi.Text = "" Then
    MsgBox "Silahkan isi lama permisi pegawai!", vbExclamation + vbOKOnly, ""
    textLamaPermisi.SetFocus
Else
    Pesan = MsgBox("Input sudah benar. Anda yakin dengan isian Anda?", vbQuestion + vbYesNo, "Konfirmasi")
    If Pesan = vbYes Then
        With AdodcUtama
            .Recordset.AddNew
            .Recordset.Fields(0).Value = textKode.Text
            .Recordset.Fields(1).Value = cmbNIP.Text
            .Recordset.Fields(2).Value = TextNama.Text
            .Recordset.Fields(3).Value = textBagian.Text
            .Recordset.Fields(4).Value = textJabatan.Text
            .Recordset.Fields(5).Value = textTanggal.Text
            .Recordset.Fields(6).Value = textBulan.Text
            .Recordset.Fields(7).Value = textTahun.Text
            .Recordset.Fields(8).Value = textKeluhan.Text
            .Recordset.Fields(9).Value = textJamPermisi.Text
            .Recordset.Fields(10).Value = textLamaPermisi.Text
            .Recordset.Fields(11).Value = cmbSatuanLamaPermisi.Text
            .Recordset.Update
            .Refresh
        End With
    cmBaru.Enabled = True
    cmSimpan.Enabled = False
    cmReset.Enabled = False
    KosongkanInput
    NonAktifkanInput
    End If
End If
Exit Sub
HancurkanError:
    PusatError
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
    NonAktifkanInput
    cmSimpan.Enabled = False
    cmReset.Enabled = False
    textJamPermisi.Text = Time
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub menuVIK_Click()
For Each Objek In Me
    If TypeName(Objek) = "TextBox" Then
        If Objek.Text = "" Then Objek.Text = "-"
    End If
Next
End Sub

