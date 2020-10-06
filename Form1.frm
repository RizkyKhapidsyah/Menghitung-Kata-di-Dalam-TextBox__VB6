VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menghitung Kata di Dalam TextBox"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  'Ketikkan beberapa buah kalimat yang cukup panjang
  'sehingga mengandung sampai puluhan bahkan ratusan
  'kata untuk mencoba fungsi menghitung kata di bawah.
  MsgBox GetWordCount(Text1.Text)
End Sub

Public Function GetWordCount(ByVal Text As String) _
As Long
    'Definisikan sebuah tanda hubung pada setiap akhir
'baris yang merupakan bagian dari seluruh kata,
'jadi kombinasikan bersama.
    Text = Trim(Replace(Text, "-" & vbNewLine, ""))
    'Ganti baris baru dengan sebuah space tunggal
    Text = Trim(Replace(Text, vbNewLine, " "))
    'Ganti spasi yang lebih dari satu (jika ada)
    'menjadi spasi tunggal
    Do While Text Like "*  *"
        Text = Replace(Text, "  ", " ")
    Loop
    'Pisahkan string dan kembalikan kata yang dihitung
    GetWordCount = 1 + UBound(Split(Text, " "))
End Function


