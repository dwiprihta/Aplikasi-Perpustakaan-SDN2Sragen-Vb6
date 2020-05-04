VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form trans_kembali 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI PENGEMBALIAN BUKA"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8940
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   20415
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         Caption         =   "DATA PENGEMBALIAN"
         BeginProperty Font 
            Name            =   "Multicolore "
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   360
         TabIndex        =   20
         Top             =   240
         Width           =   5415
      End
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      DataField       =   "TanggalKembali"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   12720
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   222232577
      CurrentDate     =   43646
   End
   Begin VB.TextBox Text26 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   360
      TabIndex        =   15
      Top             =   2880
      Width           =   6495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   360
      TabIndex        =   10
      Top             =   1440
      Width           =   8175
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   6120
         TabIndex        =   17
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   222232577
         CurrentDate     =   43646
      End
      Begin MSComCtl2.DTPicker DTPinjam 
         DataField       =   "TanggalPinjam"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   222232577
         CurrentDate     =   43646
      End
      Begin VB.TextBox Text2 
         DataField       =   "Nama"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Pinjam"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Kembali"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
   End
   Begin Crystal.CrystalReport crLAP3 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   17640
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   222298113
      CurrentDate     =   43059
   End
   Begin VB.TextBox Text5 
      DataField       =   "JudulBuku"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   17880
      TabIndex        =   8
      Text            =   "Text5"
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      DataField       =   "Alamat"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   17880
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text3 
      DataField       =   "Kelas"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   17640
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      DataField       =   "NAP"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   17760
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H8000000D&
      Caption         =   "CETAK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   195
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   14400
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\SOFTWARE PERPUSTAKAAN\perpus.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\SOFTWARE PERPUSTAKAAN\perpus.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "PEMINJAMAN"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "CARI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1530
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808000&
      Caption         =   "HAPUS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8400
      Width           =   1530
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "KEMBALIKAN BUKU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8400
      Width           =   2610
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "kembali.frx":0000
      Height          =   4575
      Left            =   360
      TabIndex        =   0
      Top             =   3600
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   -2147483644
      HeadLines       =   2
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "trans_kembali"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FORM TRANSAKSI PENGEMBALIAN
'MENAMPILKAN DATA PENGEMBALIAN BUKU DAN OPERASI (SIMPAN, HAPUS, UBAH)
'by INDRI DWI S
'======================================================================



'CARI DATA
'jika tombol cari diklik cari data peminjaman untuk dikembalikan
Private Sub Command1_Click()
If Text1.Text = "" Then
'jika data peminjaman kosong tampilkan pesan peringatan
MsgBox "Data peminjam sudah kosong!", vbCritical, "WARNING !"
Else
  'tampilkan data dari datagrid peminjaman
    If MsgBox("Apakah anda yakin ingin tetap melanjutkan pengembalian buku ini?", vbYesNo + vbDefaultButton2 + vbQuestion, "Peringatan!") = vbYes Then
    trans_selesai.Adodc1.Recordset.AddNew
    trans_selesai.Adodc1.Recordset.Fields("NAP") = Text1.Text
    trans_selesai.Adodc1.Recordset.Fields("Nama") = Text2.Text
    trans_selesai.Adodc1.Recordset.Fields("Kelas") = TEXT3.Text
    trans_selesai.Adodc1.Recordset.Fields("Alamat") = Text4.Text
    trans_selesai.Adodc1.Recordset.Fields("JudulBuku") = Text5.Text
    trans_selesai.Label8 = DTPinjam.Value
    trans_selesai.Label9 = DTPicker2.Value
    trans_selesai.Adodc1.Recordset.Update
    'jika data berhasil dikembalikan tampilkan notif suksea]s
    MsgBox "Buku berhasil dikembalikan!", vbOKOnly, "Informasi!"
    'hapus data dari data pinjam dan pindahkan ke transaksi selesai
    Adodc1.Recordset.Delete
    trans_selesai.Text2 = Text2.Text
    trans_selesai.TEXT3 = TEXT3.Text
    trans_selesai.Text4 = Text4.Text
    trans_selesai.Text5 = Text5.Text
    Else
End If
End If
End Sub

'HAPUS
'jika tombol hapus diklik
Private Sub Command2_Click()
If Text1.Text = "" Then
MsgBox "Data peminjam sudah kosong!", vbCritical, "WARNING !"
Else
        'hapus data buku belum kembali
        Dim pesan  As Integer
        pesan = MsgBox("Apakah Anda yakin ingin menghapus data ini?", vbCritical + vbYesNo, "WARNING !")
        If pesan = vbYes Then
        Adodc1.Recordset.Delete
        Else
        
End If
End If
End Sub

'PENCARIAN DATA PINJAM
'digunakan untuk mencari data buku belum kembali
Private Sub Command3_Click()
If Text26.Text = "" Then
MsgBox "ISIKAN DATA PENCARIAN ANDA!", vbOKOnly, "Informasi!"
Else
'saring data berdasarkan nama atau no perpustakaan
Adodc1.Recordset.Filter = "Nama like '%" + Me.Text26.Text + "%' or NAP like '%" + Me.Text26.Text + "%'"
End If
End Sub

'TAMPILKAN DATA LAPORAN BUKU BELUM DIKEMBALIKAN
Private Sub Command6_Click()
xx = "\LAP3.rpt"
cc = "*"
With crLAP3
 
    .ReportFileName = App.Path & xx
    .WindowState = crptMaximized
    .RetrieveDataFiles
    .Action = 1
End With
End Sub

'pindah data dari datagrid ke form
Private Sub DataGrid1_Click()
'Kode menghitung denda saaat datagrid diklik

End Sub


'KODING REFRESH DATABASE OTOMATIS SAAT FORM DIBUKA
Private Sub Form_Load()
With DataGrid1
.Columns(0).Width = 1800
.Columns(1).Width = 3600
.Columns(2).Width = 950
.Columns(3).Width = 3400
.Columns(4).Width = 4200
.Columns(5).Width = 1800


.Columns(0).Caption = "NO ANGGOTA"
.Columns(1).Caption = "NAMA ANGGOTA"
.Columns(2).Caption = "KELAS/STATUS"
.Columns(3).Caption = "ALAMAT"
.Columns(4).Caption = "JUDUL BUKU"
.Columns(5).Caption = "TANGGAL PINJAM"

.Columns(5).NumberFormat = "dd MMMM yy"
.Columns(6).NumberFormat = "dd MMMM yy"
End With
End Sub

'KODING REFRESH DATA SETELAH PENCARIAN DATA SELESAI
Private Sub Text26_Change()
If Text26.Text = "" Then
Adodc1.Refresh
With DataGrid1
.Columns(0).Width = 1800
.Columns(1).Width = 3600
.Columns(2).Width = 950
.Columns(3).Width = 3400
.Columns(4).Width = 4200
.Columns(5).Width = 1800

.Columns(0).Caption = "NO ANGGOTA"
.Columns(1).Caption = "NAMA ANGGOTA"
.Columns(2).Caption = "KELAS/STATUS"
.Columns(3).Caption = "ALAMAT"
.Columns(4).Caption = "JUDUL BUKU"
.Columns(5).Caption = "TANGGAL PINJAM"
End With
Else
'wkwk
End If
End Sub


