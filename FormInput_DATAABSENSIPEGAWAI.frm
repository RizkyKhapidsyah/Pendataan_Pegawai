VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormInput_DATAABSENSIPEGAWAI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Data Absensi"
   ClientHeight    =   5070
   ClientLeft      =   1980
   ClientTop       =   1395
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "Sylfaen"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormInput_DATAABSENSIPEGAWAI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbJenisPegawai 
      Height          =   390
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox textKode 
      Height          =   390
      Left            =   5280
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdodcUtama 
      Height          =   375
      Left            =   5280
      Top             =   3240
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
   Begin VB.CommandButton cmTutup 
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   5160
      TabIndex        =   32
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmManage 
      Caption         =   "&Manage"
      Height          =   495
      Left            =   5160
      TabIndex        =   31
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmReset 
      Caption         =   "&Reset"
      Height          =   495
      Left            =   5160
      TabIndex        =   30
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   5160
      TabIndex        =   29
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmBaru 
      Caption         =   "&Baru"
      Height          =   495
      Left            =   5160
      TabIndex        =   28
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   4935
      Begin VB.ComboBox cmbNIP 
         Height          =   390
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox textTanggal 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   2040
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox textBulan 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   2760
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox textTahun 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   3480
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cmbStatusKehadiran 
         Height          =   390
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox textKeterangan 
         Height          =   390
         Left            =   2040
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIP"
         Height          =   270
         Left            =   1470
         TabIndex        =   27
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   26
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   270
         Left            =   1185
         TabIndex        =   25
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   24
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Kehadiran"
         Height          =   270
         Left            =   465
         TabIndex        =   23
         Top             =   1200
         Width           =   1350
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   22
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         Height          =   270
         Left            =   885
         TabIndex        =   21
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   20
         Top             =   1680
         Width           =   45
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   4935
      Begin VB.TextBox textGolonganPegawai 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   390
         Left            =   2040
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox textJabatanPegawai 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   390
         Left            =   2040
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox textBagianPegawai 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   390
         Left            =   2040
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox textNamaPegawai 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   390
         Left            =   2040
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   12
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Golongan Pegawai"
         Height          =   270
         Left            =   315
         TabIndex        =   11
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   9
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan Pegawai"
         Height          =   270
         Left            =   495
         TabIndex        =   8
         Top             =   1200
         Width           =   1320
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   6
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bagian Pegawai"
         Height          =   270
         Left            =   555
         TabIndex        =   5
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pegawai"
         Height          =   270
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc AdodcNIP 
      Height          =   375
      Left            =   5280
      Top             =   3480
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
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Pegawai  : "
      Height          =   270
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   1260
   End
   Begin VB.Menu menuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu menuVIK 
         Caption         =   "Verifikasi Input Kosong"
      End
   End
End
Attribute VB_Name = "FormInput_DATAABSENSIPEGAWAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
Nyambungg
With AdodcUtama
    .ConnectionString = CN.ConnectionString
    .RecordSource = "Select * from TbAbsensiPegawai"
    .Refresh
End With
With AdodcNIP
    .ConnectionString = CN.ConnectionString
    .RecordSource = "Select * from TbDataIdentitasPegawai"
    .Refresh
End With
KosongkanInput
textKode.Text = Second(Time) & "AP" & Hour(Time) & "R" & Minute(Time)
    cmbNIP.Clear
    cmbNIP.AddItem "-"
    Do Until AdodcNIP.Recordset.EOF
        cmbNIP.AddItem AdodcNIP.Recordset.Fields(0).Value, 0
        AdodcNIP.Recordset.MoveNext
    Loop
    cmbNIP.ListIndex = 0
    textTanggal.Text = Day(Date)
    textBulan.Text = Month(Date)
    textTahun.Text = Year(Date)
    With cmbStatusKehadiran
        .Clear
        .AddItem "Hadir", 0
        .AddItem "Izin", 1
        .AddItem "Permisi", 2
        .AddItem "Alpa (Tanpa Keterangan)", 3
        .ListIndex = 0
    End With
    textNamaPegawai.Locked = True
    textJabatanPegawai.Locked = True
    textGolonganPegawai.Locked = True
    textBagianPegawai.Locked = True
    With cmbJenisPegawai
        .AddItem "Tetap", 0
        .AddItem "Honor", 1
        .AddItem "Kontrak", 2
        .ListIndex = 0
    End With
cmBaru.Enabled = False
For Each Objek In Me
    If TypeName(Objek) = "TextBox" Then Objek.MaxLength = 254
Next
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
    ElseIf TypeName(Objek) = "CommandButton" Then
        With Objek
            .Enabled = True
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
cmManage.Enabled = True
cmTutup.Enabled = True
    menuVIK.Enabled = False
End Sub


Private Sub cmBaru_Click()
    AktifkanInput
    cmBaru.Enabled = False
    cmSimpan.Enabled = True
    cmReset.Enabled = True
    KosongkanInput
    textKode.Text = Second(Time) & "AP" & Hour(Time) & "R" & Minute(Time)
    textTanggal.Text = Day(Date)
    textBulan.Text = Month(Date)
    textTahun.Text = Year(Date)
    cmbNIP.SetFocus
End Sub

Private Sub cmBaru_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub cmbJenisPegawai_Click()
If cmbJenisPegawai.ListIndex = 1 Then
    With cmbNIP
        .Text = "-"
        .Enabled = False
        .BackColor = Me.BackColor
    End With
Else
    With cmbNIP
        .Text = AdodcNIP.Recordset.Fields(0).Value
        .Enabled = True
        .BackColor = vbWhite
    End With
End If
End Sub

Private Sub cmbNIP_Click()
    AdodcNIP.Recordset.MoveFirst
    If RS.State = 1 Then RS.Close
    RS.Open "select * from tbdataidentitaspegawai where NIP = '" & cmbNIP.Text & "'", CN, 3, 3
    If Not RS.EOF Then
    textNamaPegawai.Text = RS.Fields("Nama_Pegawai")
    textBagianPegawai.Text = RS.Fields("Bagian")
    textJabatanPegawai.Text = RS.Fields("Jabatan")
    textGolonganPegawai.Text = RS.Fields("Golongan")
    End If
End Sub


Private Sub cmManage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub cmReset_Click()
    KosongkanInput
    textKode.Text = Second(Time) & "AP" & Hour(Time) & "R" & Minute(Time)
    textTanggal.Text = Day(Date)
    textBulan.Text = Month(Date)
    textTahun.Text = Year(Date)
    cmbNIP.SetFocus
End Sub

Private Sub cmReset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub cmSimpan_Click()
On Error GoTo HancurkanError
If textTanggal.Text = "" Then
    MsgBox "Silahkan isi tanggal absen!", vbExclamation + vbOKOnly, ""
    textTanggal.SetFocus
ElseIf textBulan.Text = "" Then
    MsgBox "Silahkan isi Bulan absen", vbExclamation + vbOKOnly, ""
    textBulan.SetFocus
ElseIf textTahun.Text = "" Then
    MsgBox "Silahkan isi Tahun Absen!", vbExclamation + vbOKOnly, ""
    textTahun.SetFocus
ElseIf textKeterangan.Text = "" Then
    MsgBox "Silahkan isi keterangan Absensi!", vbExclamation + vbOKOnly, ""
    textKeterangan.SetFocus
Else
    Pesan = MsgBox("Input sudah lengkap, Apakah anda yakin dengan isian Anda?", vbQuestion + vbYesNo, "Konfirmasi")
    If Pesan = vbYes Then
            With AdodcUtama
                .Recordset.AddNew
                .Recordset.Fields(0).Value = textKode.Text
                .Recordset.Fields(1).Value = cmbNIP.Text
                .Recordset.Fields(2).Value = textTanggal.Text
                .Recordset.Fields(3).Value = textBulan.Text
                .Recordset.Fields(4).Value = textTahun.Text
                .Recordset.Fields(5).Value = cmbStatusKehadiran.Text
                .Recordset.Fields(6).Value = textKeterangan.Text
                .Recordset.Fields(7).Value = textNamaPegawai.Text
                .Recordset.Fields(8).Value = textBagianPegawai.Text
                .Recordset.Fields(9).Value = textJabatanPegawai.Text
                .Recordset.Fields(10).Value = textGolonganPegawai.Text
                .Recordset.Update
                .Refresh
            End With
        KosongkanInput
        NonAktifkanInput
        cmBaru.Enabled = True
        cmSimpan.Enabled = False
        cmReset.Enabled = False
    End If
End If
Exit Sub
HancurkanError:
    PusatError
End Sub

Private Sub cmSimpan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub cmTutup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub menuVIK_Click()
For Each Objek In Me
    If TypeName(Objek) = "TextBox" Then
        If Objek.Text = "" Then Objek.Text = "-"
    End If
Next
End Sub
