VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormTambahList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tambah"
   ClientHeight    =   990
   ClientLeft      =   645
   ClientTop       =   2235
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Sylfaen"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormTambahList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   5520
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4200
      Top             =   120
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
      Caption         =   "&Batal"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox textTambahList 
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label LabelTambahList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "FormTambahList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    Nyambungg
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmOK_Click()
    If textTambahList.Text = "" Then
        MsgBox "Silahkan isi nama yang ingin dimasukkan!", vbExclamation + vbOKOnly, "Kosong?"
    Else
        With Adodc1
            .Recordset.AddNew
            .Recordset.Fields(0).Value = textTambahList.Text
            .Recordset.Update
            .Refresh
        End With
        Select Case Me.Caption
        Case "Tambah List Bagian"
            With FormInput_DATAIDENTITASPEGAWAI.AdodcBagian
                .ConnectionString = CN.ConnectionString
                .RecordSource = "Select * from tblistbagian"
                .Refresh
            End With
            With FormInput_DATAIDENTITASPEGAWAI
                .cmbBagian.Clear
                Do Until .AdodcBagian.Recordset.EOF
                    .cmbBagian.AddItem .AdodcBagian.Recordset.Fields(0).Value
                    .AdodcBagian.Recordset.MoveNext
                Loop
                .cmbBagian.Text = textTambahList.Text
            End With
        Case "Tambah List Jabatan"
            With FormInput_DATAIDENTITASPEGAWAI.AdodcJabatan
                .ConnectionString = CN.ConnectionString
                .RecordSource = "Select * from tblistjabatan"
                .Refresh
            End With
            With FormInput_DATAIDENTITASPEGAWAI
                .cmbJabatan.Clear
                Do Until .AdodcJabatan.Recordset.EOF
                    .cmbJabatan.AddItem .AdodcJabatan.Recordset.Fields(0).Value
                    .AdodcJabatan.Recordset.MoveNext
                Loop
                .cmbJabatan.Text = textTambahList.Text
            End With
        Case "Tambah List Golongan"
            With FormInput_DATAIDENTITASPEGAWAI.AdodcGolongan
                .ConnectionString = CN.ConnectionString
                .RecordSource = "Select * from tblistgolongan"
                .Refresh
            End With
            With FormInput_DATAIDENTITASPEGAWAI
            .cmbGolongan.Clear
                Do Until .AdodcGolongan.Recordset.EOF
                    .cmbGolongan.AddItem .AdodcGolongan.Recordset.Fields(0).Value
                    .AdodcGolongan.Recordset.MoveNext
                Loop
                .cmbGolongan.Text = textTambahList.Text
            End With
        Case "Tambah List Agama"
            With FormInput_DATAIDENTITASPEGAWAI.AdodcAgama
                .ConnectionString = CN.ConnectionString
                .RecordSource = "Select * from tblistagama"
                .Refresh
            End With
            With FormInput_DATAIDENTITASPEGAWAI
            .cmbAgama.Clear
                Do Until .AdodcAgama.Recordset.EOF
                    .cmbAgama.AddItem .AdodcAgama.Recordset.Fields(0).Value
                    .AdodcAgama.Recordset.MoveNext
                Loop
                .cmbAgama.Text = textTambahList.Text
            End With
        Case "Tambah List Pendidikan"
            With FormInput_DATAIDENTITASPEGAWAI.AdodcPendidikan
                .ConnectionString = CN.ConnectionString
                .RecordSource = "Select * from tblistpendidikan"
                .Refresh
            End With
            With FormInput_DATAIDENTITASPEGAWAI
            .cmbPendidikan.Clear
                Do Until .AdodcPendidikan.Recordset.EOF
                    .cmbPendidikan.AddItem .AdodcPendidikan.Recordset.Fields(0).Value
                    .AdodcPendidikan.Recordset.MoveNext
                Loop
                .cmbPendidikan.Text = textTambahList.Text
            End With
        End Select
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
