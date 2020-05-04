VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form trans_pinjam 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI PINJAM BUKU"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7275
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
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
         Caption         =   "DATA pinjam"
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
         Width           =   4215
      End
   End
   Begin VB.CommandButton Commandtambah 
      BackColor       =   &H00808000&
      Caption         =   "Tambah"
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6600
      Width           =   4635
   End
   Begin VB.CommandButton Commandselesai 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Simpan"
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6600
      Width           =   4665
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Top             =   5760
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   222298113
      CurrentDate     =   43646
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   5160
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   222298113
      CurrentDate     =   43646
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cari 2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6945
      TabIndex        =   13
      Top             =   7920
      Width           =   90
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cari"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6705
      TabIndex        =   12
      Top             =   7920
      Width           =   90
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "pinjam.frx":0000
      Height          =   375
      Left            =   9000
      TabIndex        =   11
      Top             =   6720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8880
      Top             =   6360
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
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
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
      Height          =   525
      Left            =   1920
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   4440
      Width           =   4695
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
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
      Height          =   525
      Left            =   1920
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   3720
      Width           =   4710
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
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
      Height          =   525
      Left            =   1920
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   3000
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
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
      Height          =   525
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Height          =   525
      Left            =   1920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1560
      Width           =   4695
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   8880
      Top             =   5880
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
      RecordSource    =   "ANGGOTAPERPUSTAKAAN"
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Kembali"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   14
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Pinjam   "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Judul Buku         "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat          "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Kelas/ Status           "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama        "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nomor Anggota"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   2415
   End
End
Attribute VB_Name = "trans_pinjam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FORM TRANSAKSI PINJAM
'MENAMPILKAN FORM TRANSAKSI PINJAM BUKU
'by IMAM NASUHA
'======================================================================

'variabel utuk membersihkan form
Sub bersih()
Text1.Text = ""
Text2.Text = ""
TEXT3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

'ENABEL
'membuat variabel untuk membuat form menjadi hidup (dipanggil pada tombol tambah)
Sub enabel()
Text1.Enabled = True
Text2.Enabled = True
TEXT3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text1.SetFocus
End Sub

'jika tombol cari diklik tampilkan data anggota
Private Sub Command1_Click()
cari_anggota.Show
End Sub

Private Sub Command2_Click()
cari_buku.Show
End Sub

'variabel yanag dijalankan otomastis saat form dibuka
Private Sub Form_Load()
'setting range databse
DTPicker1.Value = Now
DTPicker2.Value = DTPicker1.Value + 7

Text1.Text = "Masukkan Nomor Anggota"
Text2.Text = "Masukkan Nama"
TEXT3.Text = "Masukkan Kelas"
Text4.Text = "Masukkan Alamat"
Text5.Text = "Masukkan Judul Buku yang dipinjam"

Text1.Text = "Masukkan Nomor Anggota"
    Text2.Text = "Masukkan Nama"
    TEXT3.Text = "Masukkan Kelas"
    Text4.Text = "Masukkan Alamat"
    Text5.Text = "Masukkan Judul Buku yang dipinjam"
End Sub

'TAMBAH
'jika tombol tambah diklik
Private Sub Commandtambah_Click()
'panggil variabel membersihkan form
Call bersih
Call enabel
Command1.Enabled = True
Command2.Enabled = True
Commandtambah.Visible = False
Commandselesai.Visible = True
End Sub

'SIMPAN
'jika tombol simpan diklik simpan transaksi pinjam
Private Sub Commandselesai_Click()
'jika inputan kosong tampilkan notifikasi
If Text1 = "" Or Text2 = "" Or TEXT3 = "" Or Text4 = "" Or Text5 = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN ANDA INPUTKAN !", vbInformation, "PERHATIAN !"
Else
'proses simpan data
Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields("NAP") = Text1.Text
    Adodc1.Recordset.Fields("Nama") = Text2.Text
    Adodc1.Recordset.Fields("Kelas") = TEXT3.Text
    Adodc1.Recordset.Fields("Alamat") = Text4.Text
    Adodc1.Recordset.Fields("JudulBuku") = Text5.Text
    Adodc1.Recordset!TanggalPinjam = DTPicker1
    Adodc1.Recordset!TanggalKembali = DTPicker2
Adodc1.Recordset.Update
'tampilkan pesan sukses jika transaksi berhasil diinput
MsgBox "Data berhasil disimpan!", vbOKOnly, "Informasi!"
'panggil variabel pembersih form
Call bersih
Call enabel
End If
End Sub





