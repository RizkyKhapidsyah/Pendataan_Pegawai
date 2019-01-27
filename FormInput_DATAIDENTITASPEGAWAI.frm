VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form FormInput_DATAIDENTITASPEGAWAI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Data Identitas"
   ClientHeight    =   4815
   ClientLeft      =   405
   ClientTop       =   735
   ClientWidth     =   6705
   BeginProperty Font 
      Name            =   "Sylfaen"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormInput_DATAIDENTITASPEGAWAI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6705
   Begin MSAdodcLib.Adodc AdodcUtama 
      Height          =   330
      Left            =   120
      Top             =   6720
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
   Begin VB.CommandButton cmTutup 
      Caption         =   "&Tutup"
      Height          =   975
      Left            =   5400
      TabIndex        =   29
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmManage 
      Caption         =   "&Manage"
      Height          =   975
      Left            =   5400
      TabIndex        =   28
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmSimpan 
      Caption         =   "&Simpan"
      Height          =   975
      Left            =   5400
      TabIndex        =   27
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmBaru 
      Caption         =   "&Baru"
      Height          =   975
      Left            =   5400
      TabIndex        =   26
      Top             =   480
      Width           =   1215
   End
   Begin TabDlg.SSTab TabAja 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8070
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Biodata Diri"
      TabPicture(0)   =   "FormInput_DATAIDENTITASPEGAWAI.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Keterangan"
      TabPicture(1)   =   "FormInput_DATAIDENTITASPEGAWAI.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   30
         Top             =   720
         Width           =   4935
         Begin VB.CommandButton cmSetTanggalLahir 
            Caption         =   "&Set"
            Height          =   375
            Left            =   4320
            TabIndex        =   53
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox textStatusHubungan 
            Height          =   390
            Left            =   2040
            TabIndex        =   52
            Text            =   "Text1"
            Top             =   2640
            Width           =   2775
         End
         Begin VB.TextBox textTahunLahir 
            Alignment       =   2  'Center
            Height          =   390
            Left            =   3240
            TabIndex        =   51
            Text            =   "Text1"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox textBulanLahir 
            Alignment       =   2  'Center
            Height          =   390
            Left            =   2640
            TabIndex        =   50
            Text            =   "Text1"
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox textTanggalLahir 
            Alignment       =   2  'Center
            Height          =   390
            Left            =   2040
            TabIndex        =   47
            Text            =   "Text1"
            Top             =   1200
            Width           =   495
         End
         Begin VB.ComboBox cmbJenisKelamin 
            Height          =   390
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton cmTambahPendidikan 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4320
            TabIndex        =   35
            Top             =   2160
            Width           =   495
         End
         Begin VB.CommandButton cmTambahAgama 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4320
            TabIndex        =   34
            Top             =   1680
            Width           =   495
         End
         Begin VB.ComboBox cmbPendidikan 
            Height          =   390
            Left            =   2040
            TabIndex        =   33
            Text            =   "cmbPendidikan"
            Top             =   2160
            Width           =   2175
         End
         Begin VB.ComboBox cmbAgama 
            Height          =   390
            Left            =   2040
            TabIndex        =   32
            Text            =   "cmbAgama"
            Top             =   1680
            Width           =   2175
         End
         Begin VB.TextBox textTempatLahir 
            Height          =   390
            Left            =   2040
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   720
            Width           =   2775
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   1920
            TabIndex        =   49
            Top             =   1200
            Width           =   45
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Lahir"
            Height          =   270
            Left            =   720
            TabIndex        =   48
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   1920
            TabIndex        =   45
            Top             =   2160
            Width           =   45
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pendidikan"
            Height          =   270
            Left            =   900
            TabIndex        =   44
            Top             =   2160
            Width           =   915
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   1920
            TabIndex        =   43
            Top             =   1680
            Width           =   45
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agama"
            Height          =   270
            Left            =   1245
            TabIndex        =   42
            Top             =   1680
            Width           =   570
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   1920
            TabIndex        =   41
            Top             =   720
            Width           =   45
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tempat Lahir"
            Height          =   270
            Left            =   720
            TabIndex        =   40
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   1920
            TabIndex        =   39
            Top             =   240
            Width           =   45
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis Kelamin"
            Height          =   270
            Left            =   735
            TabIndex        =   38
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   1920
            TabIndex        =   37
            Top             =   2640
            Width           =   45
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status Hubungan"
            Height          =   270
            Left            =   450
            TabIndex        =   36
            Top             =   2640
            Width           =   1365
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4095
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4935
         Begin VB.TextBox textAlamat 
            Height          =   855
            Left            =   2040
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Text            =   "FormInput_DATAIDENTITASPEGAWAI.frx":0044
            Top             =   3120
            Width           =   2775
         End
         Begin VB.TextBox textNIP 
            Height          =   390
            Left            =   2040
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox textNamaPegawai 
            Height          =   390
            Left            =   2040
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   720
            Width           =   2775
         End
         Begin VB.ComboBox cmbBagian 
            Height          =   390
            Left            =   2040
            TabIndex        =   8
            Text            =   "cmbBagian"
            Top             =   1200
            Width           =   2175
         End
         Begin VB.ComboBox cmbJabatan 
            Height          =   390
            Left            =   2040
            TabIndex        =   7
            Text            =   "cmbJabatan"
            Top             =   1680
            Width           =   2175
         End
         Begin VB.ComboBox cmbJenisPegawai 
            Height          =   390
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2160
            Width           =   2175
         End
         Begin VB.ComboBox cmbGolongan 
            Height          =   390
            Left            =   2040
            TabIndex        =   5
            Text            =   "cmbGolongan"
            Top             =   2640
            Width           =   2175
         End
         Begin VB.CommandButton cmTambahBagian 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4320
            TabIndex        =   4
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton cmTambahJabatan 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4320
            TabIndex        =   3
            Top             =   1680
            Width           =   495
         End
         Begin VB.CommandButton cmTambahGolongan 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4320
            TabIndex        =   2
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            Height          =   270
            Left            =   120
            TabIndex        =   24
            Top             =   3120
            Width           =   1700
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   1920
            TabIndex        =   23
            Top             =   3120
            Width           =   45
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NIP"
            Height          =   270
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1700
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   1920
            TabIndex        =   21
            Top             =   240
            Width           =   45
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Pegawai"
            Height          =   270
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   1700
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   1920
            TabIndex        =   19
            Top             =   720
            Width           =   45
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bagian"
            Height          =   270
            Left            =   120
            TabIndex        =   18
            Top             =   1200
            Width           =   1700
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   1920
            TabIndex        =   17
            Top             =   1200
            Width           =   45
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jabatan"
            Height          =   270
            Left            =   120
            TabIndex        =   16
            Top             =   1680
            Width           =   1700
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   1920
            TabIndex        =   15
            Top             =   1680
            Width           =   45
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis Pegawai"
            Height          =   270
            Left            =   120
            TabIndex        =   14
            Top             =   2160
            Width           =   1700
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   1920
            TabIndex        =   13
            Top             =   2160
            Width           =   45
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Golongan"
            Height          =   270
            Left            =   120
            TabIndex        =   12
            Top             =   2640
            Width           =   1700
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   1920
            TabIndex        =   11
            Top             =   2640
            Width           =   45
         End
      End
   End
   Begin MSAdodcLib.Adodc AdodcAgama 
      Height          =   330
      Left            =   120
      Top             =   4920
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
   Begin MSAdodcLib.Adodc AdodcPendidikan 
      Height          =   330
      Left            =   120
      Top             =   5280
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
   Begin MSAdodcLib.Adodc AdodcBagian 
      Height          =   330
      Left            =   120
      Top             =   5640
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
   Begin MSAdodcLib.Adodc AdodcJabatan 
      Height          =   330
      Left            =   120
      Top             =   6000
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
   Begin MSAdodcLib.Adodc AdodcGolongan 
      Height          =   330
      Left            =   120
      Top             =   6360
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
   Begin VB.Menu MenuMenu 
      Caption         =   "Menu"
      Begin VB.Menu menuVIK 
         Caption         =   "Verifikasi Input Kosong"
      End
      Begin VB.Menu menuKI 
         Caption         =   "Kosongkan Input"
      End
   End
End
Attribute VB_Name = "FormInput_DATAIDENTITASPEGAWAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    Nyambungg
    With AdodcUtama
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tbdataidentitaspegawai"
        .Refresh
    End With
    With AdodcBagian
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tblistbagian"
        .Refresh
    End With
    With AdodcJabatan
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tblistjabatan"
        .Refresh
    End With
    With AdodcGolongan
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tblistgolongan"
        .Refresh
    End With
    With AdodcAgama
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tblistagama"
        .Refresh
    End With
    With AdodcPendidikan
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tblistpendidikan"
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
    cmbBagian.Clear
    cmbJabatan.Clear
    cmbJenisPegawai.Clear
    cmbGolongan.Clear
    cmbAgama.Clear
    cmbPendidikan.Clear
    Do Until AdodcBagian.Recordset.EOF
        cmbBagian.AddItem AdodcBagian.Recordset.Fields(0).Value
        AdodcBagian.Recordset.MoveNext
    Loop
    Do Until AdodcJabatan.Recordset.EOF
        cmbJabatan.AddItem AdodcJabatan.Recordset.Fields(0).Value
        AdodcJabatan.Recordset.MoveNext
    Loop
    Do Until AdodcGolongan.Recordset.EOF
        cmbGolongan.AddItem AdodcGolongan.Recordset.Fields(0).Value
        AdodcGolongan.Recordset.MoveNext
    Loop
    Do Until AdodcAgama.Recordset.EOF
        cmbAgama.AddItem AdodcAgama.Recordset.Fields(0).Value
        AdodcAgama.Recordset.MoveNext
    Loop
    Do Until AdodcPendidikan.Recordset.EOF
        cmbPendidikan.AddItem AdodcPendidikan.Recordset.Fields(0).Value
        AdodcPendidikan.Recordset.MoveNext
    Loop
    cmbBagian.Text = "[Pilih Bagian].."
    cmbJabatan.Text = "[Pilih Jabatan].."
    cmbGolongan.Text = "[Pilih Golongan].."
    cmbAgama.Text = "[Pilih Agama].."
    cmbPendidikan.Text = "[Pilih Pendidikan].."
    With cmbJenisPegawai
        .AddItem "Tetap", 0
        .AddItem "Honor", 1
        .AddItem "Kontrak", 2
        .ListIndex = 0
    End With
    With cmbJenisKelamin
        .Clear
        .AddItem "Laki-Laki", 0
        .AddItem "Perempuan", 1
        .ListIndex = 0
    End With
    TabAja.Tab = 0
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
    ElseIf TypeName(Objek) = "CommandButton" Then
        With Objek
            .Enabled = False
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
    KosongkanInput
    textNIP.SetFocus
    TabAja.Tab = 0
End Sub

Private Sub cmBaru_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub cmbJenisPegawai_Click()
    If cmbJenisPegawai.ListIndex = 1 Then
        With cmbGolongan
            .Text = "-"
            .Enabled = False
        End With
        cmTambahGolongan.Enabled = False
    Else
        With cmbGolongan
            .Enabled = True
            .Text = "[Pilih Golongan].."
        End With
        cmTambahGolongan.Enabled = True
    End If
End Sub

Private Sub cmManage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub cmSetTanggalLahir_Click()
    With FormKalender
        .Caption = "Set Tanggal Lahir"
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmSetTanggalLahir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub cmSimpan_Click()
On Error GoTo HancurkanError
If textNIP.Text = "" Then
    MsgBox "Silahkan isi NIP Pegawai", vbExclamation + vbOKOnly, ""
    textNIP.SetFocus
    TabAja.Tab = 0
ElseIf textNamaPegawai.Text = "" Then
    MsgBox "Silahkan isi Nama Pegawai", vbExclamation + vbOKOnly, ""
    textNamaPegawai.SetFocus
    TabAja.Tab = 0
ElseIf cmbBagian.Text = "" Or cmbBagian.Text = "[Pilih Bagian].." Then
    MsgBox "Silahkan pilih bagian Pegawai", vbExclamation + vbOKOnly, ""
    cmbBagian.SetFocus
    TabAja.Tab = 0
ElseIf cmbJabatan.Text = "" Or cmbJabatan.Text = "[Pilih Jabatan].." Then
    MsgBox "Silahkan pilih Jabatan Pegawai", vbExclamation + vbOKOnly, ""
    cmbJabatan.SetFocus
    TabAja.Tab = 0
ElseIf cmbGolongan.Text = "" Or cmbGolongan.Text = "[Pilih Golongan].." Then
    MsgBox "Silahkan pilih Golongan Pegawai", vbExclamation + vbOKOnly, ""
    cmbGolongan.SetFocus
    TabAja.Tab = 0
ElseIf textAlamat.Text = "" Then
    MsgBox "Silahkan isi alamat pegawai!", vbExclamation + vbOKOnly, ""
    textAlamat.SetFocus
    TabAja.Tab = 0
ElseIf textTempatLahir.Text = "" Then
    MsgBox "Silahkan isi kota lahir pegawai!", vbExclamation + vbOKOnly, ""
    textTempatLahir.SetFocus
    TabAja.Tab = 1
ElseIf textTanggalLahir.Text = "" Then
    MsgBox "Silahkan isi tanggal lahir pegawai", vbExclamation + vbOKOnly, ""
    textTanggalLahir.SetFocus
    TabAja.Tab = 1
ElseIf textBulanLahir.Text = "" Then
    MsgBox "Silahkan isi bulan lahir pegawai", vbExclamation + vbOKOnly, ""
    textBulanLahir.SetFocus
    TabAja.Tab = 1
ElseIf textTahunLahir.Text = "" Then
    MsgBox "Silahkan isi tahun lahir pegawai", vbExclamation + vbOKOnly, ""
    textTahunLahir.SetFocus
    TabAja.Tab = 1
ElseIf cmbAgama.Text = "" Or cmbAgama.Text = "[Pilih Agama].." Then
    MsgBox "Silahkan isi Agama!", vbExclamation + vbOKOnly, ""
    cmbAgama.SetFocus
    TabAja.Tab = 1
ElseIf cmbPendidikan.Text = "" Or cmbPendidikan.Text = "[Pilih Pendidikan].." Then
    MsgBox "Silahkan isi Pendidikan terakhir pegawai!", vbExclamation + vbOKOnly, ""
    cmbPendidikan.SetFocus
    TabAja.Tab = 1
ElseIf textStatusHubungan.Text = "" Then
    MsgBox "silahkan isi status hubungan pegawai!", vbExclamation + vbOKOnly, ""
    textStatusHubungan.SetFocus
    TabAja.Tab = 1
Else
    Pesan = MsgBox("Anda yakin bahwa isian Anda sudah benar?", vbQuestion + vbYesNo, "Konfirmasi")
    If Pesan = vbYes Then
        With AdodcUtama
            .Recordset.AddNew
            .Recordset.Fields(0).Value = textNIP.Text
            .Recordset.Fields(1).Value = textNamaPegawai.Text
            .Recordset.Fields(2).Value = cmbBagian.Text
            .Recordset.Fields(3).Value = cmbJabatan.Text
            .Recordset.Fields(4).Value = cmbJenisPegawai.Text
            .Recordset.Fields(5).Value = cmbGolongan.Text
            .Recordset.Fields(6).Value = textAlamat.Text
            .Recordset.Fields(7).Value = cmbJenisKelamin.Text
            .Recordset.Fields(8).Value = textTempatLahir.Text
            .Recordset.Fields(9).Value = textTanggalLahir.Text & " - " & textBulanLahir.Text & " - " & textTahunLahir.Text
            .Recordset.Fields(10).Value = cmbAgama.Text
            .Recordset.Fields(11).Value = cmbPendidikan.Text
            .Recordset.Fields(12).Value = textStatusHubungan.Text
            .Recordset.Update
            .Refresh
        End With
                KosongkanInput
                NonAktifkanInput
                cmBaru.Enabled = True
                cmSimpan.Enabled = False
    End If
End If
Exit Sub
HancurkanError:
    PusatError
End Sub

Private Sub cmSimpan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub cmTambahAgama_Click()
    With FormTambahList
        .Caption = "Tambah List Agama"
        .LabelTambahList = "Masukkan Nama Agama : "
        .textTambahList.Text = ""
        .Adodc1.ConnectionString = CN.ConnectionString
        .Adodc1.RecordSource = "Select * From TbListAgama"
        .Adodc1.Refresh
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmTambahAgama_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub cmTambahBagian_Click()
    With FormTambahList
        .Caption = "Tambah List Bagian"
        .LabelTambahList = "Masukkan Nama Bagian : "
        .textTambahList.Text = ""
        .Adodc1.ConnectionString = CN.ConnectionString
        .Adodc1.RecordSource = "Select * From TbListBagian"
        .Adodc1.Refresh
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmTambahBagian_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub cmTambahGolongan_Click()
    With FormTambahList
        .Caption = "Tambah List Golongan"
        .LabelTambahList = "Masukkan Nama Golongan : "
        .textTambahList.Text = ""
        .Adodc1.ConnectionString = CN.ConnectionString
        .Adodc1.RecordSource = "Select * From TbLIstGolongan"
        .Adodc1.Refresh
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmTambahGolongan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub cmTambahJabatan_Click()
    With FormTambahList
        .Caption = "Tambah List Jabatan"
        .LabelTambahList = "Masukkan Nama Jabatan : "
        .textTambahList.Text = ""
        .Adodc1.ConnectionString = CN.ConnectionString
        .Adodc1.RecordSource = "Select * From TbLIstJabatan"
        .Adodc1.Refresh
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmTambahJabatan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub

Private Sub cmTambahPendidikan_Click()
    With FormTambahList
        .Caption = "Tambah List Pendidikan"
        .LabelTambahList = "Masukkan Nama Pendidikan : "
        .textTambahList.Text = ""
        .Adodc1.ConnectionString = CN.ConnectionString
        .Adodc1.RecordSource = "Select * From TbListPendidikan"
        .Adodc1.Refresh
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmTambahPendidikan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    AktifkanInput
    KosongkanInput
    cmBaru.Enabled = False
    cmSimpan.Enabled = True
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

Private Sub menuKI_Click()
    KosongkanInput
    TabAja.Tab = 0
    textNIP.SetFocus
End Sub

Private Sub menuVIK_Click()
For Each Objek In Me
    If TypeName(Objek) = "TextBox" Then
        If Objek.Text = "" Then Objek.Text = "-"
    ElseIf TypeName(Objek) = "ComboBox" Then
        If Objek.Text = "" Then Objek.Text = "-"
    End If
Next
End Sub

Private Sub TabAja_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MenuMenu
End Sub
