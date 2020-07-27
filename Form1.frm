VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memisahkan Komponen dari Suatu Tanggal"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   1920
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim hari As Integer, bulan As Integer, tahun As Integer
  hari = DateTime.DatePart("d", _
                   CDate("22/01/1973"), _
                   vbUseSystemDayOfWeek, _
                   vbUseSystem) 'Menghasilkan 22
    bulan = DateTime.DatePart("m", _
                   CDate("22/01/1973"), _
                   vbUseSystemDayOfWeek, _
                   vbUseSystem) 'Menghasilkan 1
    tahun = DateTime.DatePart("yyyy", _
                   CDate("22/01/1973"), _
                   vbUseSystemDayOfWeek, _
                   vbUseSystem) 'Menghasilkan 1973
  MsgBox hari
  MsgBox bulan
  MsgBox tahun

End Sub
