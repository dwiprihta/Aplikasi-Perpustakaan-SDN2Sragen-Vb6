VERSION 5.00
Begin VB.Form index 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H000080FF&
   Caption         =   "PERPUSTAKAAN SD NEGERI 2 SRAGEN"
   ClientHeight    =   10755
   ClientLeft      =   75
   ClientTop       =   720
   ClientWidth     =   20370
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "index.frx":0000
   ScaleHeight     =   10755
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3975
      Left            =   15000
      TabIndex        =   8
      Top             =   3240
      Width           =   3615
      Begin VB.Line Line4 
         BorderColor     =   &H8000000A&
         X1              =   360
         X2              =   3360
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "log-out"
         BeginProperty Font 
            Name            =   "Multicolore "
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   9
         Top             =   3360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3975
      Left            =   10560
      TabIndex        =   5
      Top             =   3240
      Width           =   3615
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PENGEMBALIAN"
         BeginProperty Font 
            Name            =   "Multicolore "
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   6
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000A&
         X1              =   360
         X2              =   3360
         Y1              =   3240
         Y2              =   3240
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3975
      Left            =   6240
      TabIndex        =   3
      Top             =   3240
      Width           =   3615
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PEMINJAMAN"
         BeginProperty Font 
            Name            =   "Multicolore "
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   7
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         X1              =   360
         X2              =   3360
         Y1              =   3240
         Y2              =   3240
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3975
      Left            =   1920
      TabIndex        =   2
      Top             =   3240
      Width           =   3615
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   360
         X2              =   3360
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DATA BUKU"
         BeginProperty Font 
            Name            =   "Multicolore "
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   4
         Top             =   3360
         Width           =   2055
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label77 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "--:--:--"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   10680
      TabIndex        =   1
      Top             =   7440
      Width           =   2895
   End
   Begin VB.Label Label88 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "--/--/----"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   7440
      Width           =   9015
   End
   Begin VB.Menu MASTER 
      Caption         =   "MASTER"
      Begin VB.Menu DABUK 
         Caption         =   "DATA BUKU"
      End
   End
   Begin VB.Menu TRANS 
      Caption         =   "TRANSAKSI"
      Begin VB.Menu pin 
         Caption         =   "PEMINJAMAN"
      End
      Begin VB.Menu PENG 
         Caption         =   "PENGEMBALIAN"
      End
   End
   Begin VB.Menu OUT 
      Caption         =   "LOG-OUT"
   End
End
Attribute VB_Name = "index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'FORM HALAMAN UTAMA APLIKASI
'MENAMPILKAN MENU KE SELURUH FORM
'by IMAM NASUHA
'======================================================================

'TAMPILKAN FORM DATA KEMBALI BUKU
Private Sub DAPENG_Click()
kembali.Show
End Sub

'TAMPILKAN FORM DATA BUKU
Private Sub DABUK_Click()
data_buku.Show
End Sub

'TAMPILKAN FORM DATA ANGGOTA
Private Sub DATAANG_Click()

End Sub

'TAMPILKAN FORM DATA TRANSAKSI
Private Sub DATRANS_Click()
trans_selesai.Show
End Sub

Private Sub dosen_Click()
anggota_guru.Show
End Sub



Private Sub Frame1_Click()
data_buku.Show
End Sub

Private Sub Frame2_Click()
trans_pinjam.Show
End Sub

Private Sub Frame3_Click()
trans_kembali.Show
End Sub

Private Sub Frame4_Click()
If MsgBox("Apakah Anda yakin ingin keluar ?", vbYesNo + vbDefaultButton2 + vbQuestion, "VB 6.0 WARNING !") = vbYes Then
End
End If
End Sub

'TAMPILKAN WAKTU
Private Sub Label2_Click()
Label2.Caption = Time
End Sub

'PERTANYAAN SAAT AKAN KELUAR
Private Sub OUT_Click()
If MsgBox("Apakah Anda yakin ingin keluar ?", vbYesNo + vbDefaultButton2 + vbQuestion, "VB 6.0 WARNING !") = vbYes Then
End
End If
End Sub

'TAMPILKAN FORM TRANSAKSI PEMINJAMAN
Private Sub PEMIN_Click()
trans_pinjam.Show
End Sub

'TAMPILKAN FORM PINJAM BUKU
Private Sub PENG_Click()
trans_kembali.Show
End Sub

'TAMPILKAN FORM TRANSAKSI PEMINJAMAN
Private Sub pin_Click()
trans_pinjam.Show
End Sub

Private Sub siswa_Click()
data_anggota.Show
End Sub

'TAMPILKAN WAKTU
Private Sub Timer1_Timer()
Label77.Caption = Format(Now, "hh : mm : ss")
'Label88.Caption = Format(Now, "dd MMMM yyyy")
Label88.Caption = Format(Now, "dd MMMM yyyy")
   'Label2.Caption = Time
End Sub




