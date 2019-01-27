VERSION 5.00
Begin VB.MDIForm FormUtama 
   BackColor       =   &H8000000C&
   Caption         =   "Aplikasi Pendataan Pegawai PT. PLN (Persero)"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   4620
      TabIndex        =   0
      Top             =   0
      Width           =   4680
   End
   Begin VB.Menu menuPegawai 
      Caption         =   "Pegawai"
      Begin VB.Menu menuIDIP 
         Caption         =   "Input Data Identitas Pegawai"
      End
      Begin VB.Menu menuIDAP 
         Caption         =   "Input Data Absensi Pegawai"
      End
      Begin VB.Menu menuIDKP 
         Caption         =   "Input Data Keluhan Pegawai"
      End
      Begin VB.Menu menuIDLP 
         Caption         =   "Input Data Lembur Pegawai"
      End
      Begin VB.Menu menuIDCP 
         Caption         =   "Input Data Cuti Pegawai"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu menuKeluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu menuManage 
      Caption         =   "Manage"
      Begin VB.Menu menuMDIP 
         Caption         =   "Manage Data Identitas Pegawai"
      End
      Begin VB.Menu menuMDAP 
         Caption         =   "Manage Data Absensi Pegawai"
      End
      Begin VB.Menu menuMDKP 
         Caption         =   "Manage Data Keluhan Pegawai"
      End
      Begin VB.Menu menuMDLP 
         Caption         =   "Manage Data Lembur Pegawai"
      End
      Begin VB.Menu MenuMDCP 
         Caption         =   "Manage Data Cuti Pegawai"
      End
   End
   Begin VB.Menu menuLaporan 
      Caption         =   "Laporan"
      Begin VB.Menu menuLDIP 
         Caption         =   "Laporan Data Identitas Pegawai"
         Begin VB.Menu menuLDIP_Keseluruhan 
            Caption         =   "Keseluruhan"
         End
         Begin VB.Menu sep2 
            Caption         =   "-"
         End
         Begin VB.Menu menuLDIP_PerHari 
            Caption         =   "Per Hari"
         End
         Begin VB.Menu menuLDIP_PerMinggu 
            Caption         =   "Per Minggu"
         End
         Begin VB.Menu menuLDIP_PerBulan 
            Caption         =   "Per Bulan"
         End
         Begin VB.Menu menuLDIP_PerTahun 
            Caption         =   "Per Tahun"
         End
      End
      Begin VB.Menu menuLDAP 
         Caption         =   "Laporan Data Absensi Pegawai"
         Begin VB.Menu menuLDAP_Keseluruhan 
            Caption         =   "Keseluruhan"
         End
         Begin VB.Menu sep4 
            Caption         =   "-"
         End
         Begin VB.Menu menuLDAP_PerHari 
            Caption         =   "Per Hari"
         End
         Begin VB.Menu menuLDAP_PerMinggu 
            Caption         =   "Per Minggu"
         End
         Begin VB.Menu menuLDAP_PerBulan 
            Caption         =   "Per Bulan"
         End
         Begin VB.Menu menuLDAP_PerTahun 
            Caption         =   "Per Tahun"
         End
      End
      Begin VB.Menu menuLDKP 
         Caption         =   "Laporan Data Keluhan Pegawai"
         Begin VB.Menu menuLDKP_Keseluruhan 
            Caption         =   "Keseluruhan"
         End
         Begin VB.Menu sep5 
            Caption         =   "-"
         End
         Begin VB.Menu menuLDKP_PerHari 
            Caption         =   "Per Hari"
         End
         Begin VB.Menu menuLDKP_PerMinggu 
            Caption         =   "Per Minggu"
         End
         Begin VB.Menu menuLDKP_PerBulan 
            Caption         =   "Per Bulan"
         End
         Begin VB.Menu menuLDKP_PerTahun 
            Caption         =   "Per Tahun"
         End
      End
      Begin VB.Menu menuLDLP 
         Caption         =   "Laporan Data Lembur Pegawai"
         Begin VB.Menu menuLDLP_Keseluruhan 
            Caption         =   "Keseluruhan"
         End
         Begin VB.Menu sep6 
            Caption         =   "-"
         End
         Begin VB.Menu menuLDLP_PerHari 
            Caption         =   "Per Hari"
         End
         Begin VB.Menu menuLDLP_PerMinggu 
            Caption         =   "Per Minggu"
         End
         Begin VB.Menu menuLDLP_PerBulan 
            Caption         =   "Per Bulan"
         End
         Begin VB.Menu menuLDLP_PerTahun 
            Caption         =   "Per Tahun"
         End
      End
      Begin VB.Menu menuLDCP 
         Caption         =   "Laporan Data Cuti Pegawai"
         Begin VB.Menu menuLDCP_Keseluruhan 
            Caption         =   "Keseluruhan"
         End
         Begin VB.Menu sep7 
            Caption         =   "-"
         End
         Begin VB.Menu menuLDCP_PerHari 
            Caption         =   "Per Hari"
         End
         Begin VB.Menu menuLDCP_PerMinggu 
            Caption         =   "Per Minggu"
         End
         Begin VB.Menu menuLDCP_PerBulan 
            Caption         =   "Per Bulan"
         End
         Begin VB.Menu menuLDCP_PerTahun 
            Caption         =   "Per Tahun"
         End
      End
   End
End
Attribute VB_Name = "FormUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    Me.WindowState = vbMaximized
End Sub

Private Sub MDIForm_Load()
    AturKontrol
End Sub

Private Sub menuIDAP_Click()
With FormInput_DATAABSENSIPEGAWAI
    .Show
    .SetFocus
End With
End Sub

Private Sub menuIDIP_Click()
    With FormInput_DATAIDENTITASPEGAWAI
        .Show
        .SetFocus
    End With
End Sub

Private Sub menuIDKP_Click()
With FormInput_DATAKELUHANPEGAWAI
    .Show
    .SetFocus
End With
End Sub

Private Sub menuIDLP_Click()
With FormInput_DATALEMBURPEGAWAI
    .Show
    .SetFocus
End With
End Sub
