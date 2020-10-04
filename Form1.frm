VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Hanya Huruf Besar Boleh Dientri ke TextBox"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Created by Rizky Khapidsyah
'Source code program dimulai dari sini

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
End Sub

Private Sub Text1_Change()   'Text1 menggunakan event
                             'Change
Dim posisi As Integer
posisi = Text1.SelStart
  Text1.Text = UCase(Text1.Text)
  Text1.SelStart = posisi
End Sub
     
Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


