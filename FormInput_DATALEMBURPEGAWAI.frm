VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormInput_DATALEMBURPEGAWAI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Data Lembur"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Sylfaen"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormInput_DATALEMBURPEGAWAI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   5790
   Begin VB.TextBox textKode 
      Height          =   390
      Left            =   120
      TabIndex        =   36
      Text            =   "Text2"
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton cmTutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   4680
      TabIndex        =   35
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmManage 
      Caption         =   "&Manage"
      Height          =   375
      Left            =   3480
      TabIndex        =   34
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   2280
      TabIndex        =   33
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   1200
      TabIndex        =   32
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmBaru 
      Caption         =   "&Baru"
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   6120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.TextBox TextTotalJamLembur 
         Height          =   390
         Left            =   1440
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox textLemburKe 
         Height          =   390
         Left            =   1440
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   2640
         Width           =   975
      End
      Begin VB.ComboBox cmbJenisPegawai 
         Height          =   390
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox textKeterangan 
         Height          =   855
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Text            =   "FormInput_DATALEMBURPEGAWAI.frx":000C
         Top             =   5040
         Width           =   2895
      End
      Begin VB.TextBox textTujuanLembur 
         Height          =   390
         Left            =   1440
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   4560
         Width           =   2895
      End
      Begin VB.CommandButton cmDeteksiTanggal 
         Caption         =   "<>"
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox TextJamMulaiLembur 
         Height          =   390
         Left            =   1440
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox textTahun 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   2880
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox textBulan 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   2160
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox textTanggal 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   1440
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox textJabatan 
         Height          =   390
         Left            =   1440
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox textBagian 
         Height          =   390
         Left            =   1440
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox TextNama 
         Height          =   390
         Left            =   1440
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1200
         Width           =   2895
      End
      Begin VB.ComboBox cmbNIP 
         Height          =   390
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Jam"
         Height          =   270
         Left            =   420
         TabIndex        =   42
         Top             =   4080
         Width           =   780
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1305
         TabIndex        =   41
         Top             =   4080
         Width           =   45
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lembur Ke"
         Height          =   270
         Left            =   285
         TabIndex        =   39
         Top             =   2640
         Width           =   915
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1305
         TabIndex        =   38
         Top             =   2640
         Width           =   45
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Pegawai"
         Height          =   270
         Left            =   135
         TabIndex        =   30
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   29
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   28
         Top             =   5040
         Width           =   45
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         Height          =   270
         Left            =   285
         TabIndex        =   27
         Top             =   5040
         Width           =   930
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   26
         Top             =   4560
         Width           =   45
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tujuan"
         Height          =   270
         Left            =   645
         TabIndex        =   25
         Top             =   4560
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   24
         Top             =   3600
         Width           =   45
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Mulai"
         Height          =   270
         Left            =   375
         TabIndex        =   23
         Top             =   3600
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   22
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   270
         Left            =   585
         TabIndex        =   21
         Top             =   3120
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   20
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         Height          =   270
         Left            =   615
         TabIndex        =   19
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   18
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bagian"
         Height          =   270
         Left            =   675
         TabIndex        =   17
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   16
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   270
         Left            =   720
         TabIndex        =   15
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1320
         TabIndex        =   14
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIP"
         Height          =   270
         Left            =   870
         TabIndex        =   13
         Top             =   720
         Width           =   345
      End
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
Attribute VB_Name = "FormInput_DATALEMBURPEGAWAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
Nyambungg
With AdodcUtama
    .ConnectionString = CN.ConnectionString
    .RecordSource = "Select * From TbLemburPegawai"
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
    textKode.Text = Second(Time) & "LP" & Hour(Time) & "R-PLN" & Minute(Time)
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
ElseIf textLemburKe.Text = "" Then
    MsgBox "Silahkan isi Data Lembur Ke Berapa pada pegawai!", vbExclamation + vbOKOnly, ""
    textLemburKe.SetFocus
ElseIf TextJamMulaiLembur.Text = "" Then
    MsgBox "Silahkan isi data jam mulai lembur pegawai!", vbExclamation + vbOKOnly, ""
    TextJamMulaiLembur.SetFocus
ElseIf TextTotalJamLembur.Text = "" Then
    MsgBox "Silahkan isi data total jam lembur pegawai!", vbExclamation + vbOKOnly, ""
    TextTotalJamLembur.SetFocus
ElseIf textTujuanLembur.Text = "" Then
    MsgBox "Silahkan isi data tujuan lembur pegawai!", vbExclamation + vbOKOnly, ""
    textTujuanLembur.SetFocus
ElseIf textKeterangan.Text = "" Then
    MsgBox "Silahkan isi data keterangan lembur pegawai!", vbExclamation + vbOKOnly, ""
    textKeterangan.SetFocus
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
            .Recordset.Fields(5).Value = textLemburKe.Text
            .Recordset.Fields(6).Value = textTanggal.Text
            .Recordset.Fields(7).Value = textBulan.Text
            .Recordset.Fields(8).Value = textTahun.Text
            .Recordset.Fields(9).Value = TextJamMulaiLembur.Text
            .Recordset.Fields(10).Value = TextTotalJamLembur.Text
            .Recordset.Fields(11).Value = textTujuanLembur.Text
            .Recordset.Fields(12).Value = textKeterangan.Text
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


