VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Random Systems"
   ClientHeight    =   7860
   ClientLeft      =   420
   ClientTop       =   1005
   ClientWidth     =   15270
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   15270
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Dolu"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   21
      Top             =   6480
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "No"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   20
      Top             =   5760
      Width           =   615
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Düz"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   7080
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":628A
      Left            =   5520
      List            =   "Form1.frx":62A3
      TabIndex        =   9
      Text            =   "1x"
      Top             =   120
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Artýþ miktarý sabit"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "Tekrarla"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   3240
      Top             =   4560
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "Dur"
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2280
      Top             =   4560
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Caption         =   "Grafik"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Baþla"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000005&
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000005&
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "kaplý alan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "max kaplý alan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bayram Kotan bayramkotan@gmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      TabIndex        =   13
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   "Hýz="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8760
      TabIndex        =   5
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "S="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8400
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim genislik, yukseklik As Integer
Dim alan_gen, alan_yuk As Integer
Dim basamak_sayisi As Integer
Dim bas_x, bas_y, son_x, son_y As Double
Dim yatay, dusey As Double
Dim f As New Form2

Private Sub Check2_Click()
Form1.Cls
izgara_ciz
ciz
Form1.Refresh
End Sub

Private Sub Check3_Click()
Form1.Cls
izgara_ciz
ciz
Form1.Refresh
End Sub

Private Sub Check4_Click()
Form1.Cls
izgara_ciz
ciz
Form1.Refresh
End Sub

Private Sub Combo1_Change()
Command3_Click
If Combo1.Text = "1x" Then
 Timer1.Interval = 500
ElseIf Combo1.Text = "2x" Then
  Timer1.Interval = 250
ElseIf Combo1.Text = "0.5x" Then
   Timer1.Interval = 1000
ElseIf Combo1.Text = "0.25x" Then
   Timer1.Interval = 2000
ElseIf Combo1.Text = "5x" Then
   Timer1.Interval = 100
ElseIf Combo1.Text = "10x" Then
   Timer1.Interval = 50
ElseIf Combo1.Text = "20x" Then
   Timer1.Interval = 25
End If
End Sub

Private Sub Combo1_Click()
If Timer1.Enabled = True Then
  Command3_Click
  Command1_Click
Else
  Combo1_Change
End If
End Sub

Private Sub Command1_Click()
'hal 300 den fazla olmasýn
If hal >= 300 Then Exit Sub
'interval ý ayarla
If Combo1.Text = "1x" Then
 Timer1.Interval = 500
ElseIf Combo1.Text = "2x" Then
  Timer1.Interval = 250
ElseIf Combo1.Text = "0.5x" Then
   Timer1.Interval = 1000
ElseIf Combo1.Text = "0.25x" Then
   Timer1.Interval = 2000
ElseIf Combo1.Text = "5x" Then
   Timer1.Interval = 100
ElseIf Combo1.Text = "10x" Then
   Timer1.Interval = 50
ElseIf Combo1.Text = "20x" Then
   Timer1.Interval = 25
End If

Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
If hal >= 2 Then
 f.Show
 Timer1.Enabled = False
End If
End Sub

Private Sub Command3_Click()
Timer1.Enabled = False
End Sub

Private Sub Command4_Click()
Form_Load
Timer1.Enabled = False

Label5.Caption = "16" 'noktalarýn baþlangýç konumunda kapladýklarý alan

'ilk deðerler
'***********************************
Label8.Caption = 1

hal = 0
kapli_alan = 16
entropi = 0
ent_hesapla
Label3.Caption = entropi
'***********************************

End Sub

Private Sub Form_Activate()
f.Cls
f.Hide
End Sub
Private Sub Form_Load()
'yuklenirken elde edilen deðerler ve yapýlamasý gerekenler
'Ekran
genislik = Screen.Width
yukseklik = Screen.Height

'alan
alan_gen = genislik * 0.88
alan_yuk = yukseklik * 0.88

'alanýn içinin baþlagýç ve bitiþ deðerleri

bas_x = genislik * 0.06
bas_y = yukseklik * 0.06
son_x = genislik * 0.94
son_y = yukseklik * 0.94


'___________basamak belirle_____________
basamak_sayisi = 20

'ýzgara aralýklarýnýn geniþlikleri

yatay = alan_gen / basamak_sayisi
dusey = alan_yuk / basamak_sayisi

'ekrana göre ayarla
Me.Height = yukseklik
Me.Width = genislik
'***********************************
'kaplý olan alan ilk baþta 16
Label5.Caption = 16
'ilk deðerler
'***********************************
max_kapli = 16
Label11.Caption = "16"
Label8.Caption = 1
hal = 0
kapli_alan = 16
ent_hesapla
Label3.Caption = entropi
Label12.Caption = "% 4"
Label13.Caption = "% 4"
'***********************************
Form1.Cls
izgara_ciz
tanecik_koy
ciz
Form1.Refresh
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Caption = "Random Systems   Alan=400 birim     X=" & X & ", Y=" & Y
End Sub

'___________IZGARA ÇÝZ__________________

Private Sub izgara_ciz()
Dim i, j As Integer

'dýþ hattý çiz
Me.Line (genislik * 0.05, yukseklik * 0.05)-(genislik * 0.95, yukseklik * 0.95), vbBlack, B
Me.Line (genislik * 0.055, yukseklik * 0.055)-(genislik * 0.945, yukseklik * 0.945), vbBlack, B
Me.Line (genislik * 0.06, yukseklik * 0.06)-(genislik * 0.94, yukseklik * 0.94), vbBlack, B

If Check2.Value = False Then
  'ýzgarayý çiz 100x100 lük
  ' n-1 er tane
  For i = 1 To basamak_sayisi - 1
    'satýrlar yani sadece y artýr
    Line (bas_x, bas_y + i * dusey)-(son_x, bas_y + i * dusey)
    'sütunlar yani sadece x artýr
    Line (bas_x + i * yatay, bas_y)-(bas_x + i * yatay, son_y)
  Next
End If

End Sub

Private Sub ciz()
Dim i  As Integer
 'FillStyle = 0
 'FillColor = vbBlack
  For i = 0 To 399
   FontSize = 4
    If Not Check3.Value = False Then
        Dim str As String
        str = CStr(i + 1)
        CurrentX = X_(i) - 90
        CurrentY = Y_(i) - 90
        Print (str)
        If Check4.Value Then Me.Circle (X_(i), Y_(i)), 200, vbRed
    Else
        If Check4.Value Then Me.Circle (X_(i), Y_(i)), 70, vbRed
    End If
  Next
End Sub

Private Sub rast() 'rastgele deðerler vermek için GEÇÝCÝ
Dim i_ As Integer
Randomize
For i = 0 To 399
    X_(i_) = genislik * Rnd(1)
    Y_(i_) = yukseklik * Rnd(1)
Next
End Sub

 'form load edilirken taneciklerin ilk konumlarýný belirle
 '16x25 yani 16 hücreye 25 þer tane tanecik koy.Tam ortaya olacak þekilde
Private Sub tanecik_koy()
Dim i, j, k As Integer
   
  For j = 0 To 3
    For k = 0 To 3
         For i = 0 To 24
              X_(i + k * 25 + j * 100) = bas_x + (8.5 + k) * yatay
         Next
    Next
  Next
  
  For j = 0 To 3
        For i = 0 To 99
            Y_(i + j * 100) = bas_y + (8.5 + j) * dusey
        Next
  Next
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
Form1.Cls
izgara_ciz
hesapla  'noktalarýn sonraki konumlarýný hesapla
ciz
Form1.Refresh

alan_hesapla 'noktalarýn kapladýklarý alaný hesapla
Label5.Caption = kapli_alan
Label8.Caption = hal
Label12.Caption = "%" & oran1
Label13.Caption = "%" & oran2
If hal >= 300 Then Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Label1.Caption = CStr(Time)
End Sub

Private Sub hesapla()
 Dim i, j As Integer
Dim x_yon, y_yon As Double
Dim x_artis, y_artis As Double

Randomize

For i = 0 To 399

 x_yon = Rnd(1)
 y_yon = Rnd(1)
 
'artýþ miktarýnýn sabit mi yoksa rastgele mi olduðunu belirler
If Check1.Value = False Then
 x_artis = Rnd(1) * yatay
 y_artis = Rnd(1) * dusey
Else
 x_artis = yatay
 y_artis = dusey
End If

 'x deki artýþ
 If x_yon > 0.5 Then
   X_(i) = X_(i) + x_artis
 ElseIf x_yon < 0.5 Then
   X_(i) = X_(i) - x_artis
 End If
 
 If X_(i) >= son_x Then
  X_(i) = X_(i) - 2 * x_artis
 ElseIf X_(i) <= bas_x Then
   X_(i) = X_(i) + 2 * x_artis
 End If
   
 'y deki artýþ
  If y_yon > 0.5 Then
   Y_(i) = Y_(i) + y_artis
 ElseIf y_yon < 0.5 Then
   Y_(i) = Y_(i) - y_artis
 End If
 
 If Y_(i) >= son_y Then
  Y_(i) = Y_(i) - 2 * y_artis
 ElseIf Y_(i) <= bas_y Then
   Y_(i) = Y_(i) + 2 * y_artis
 End If
 
Next
End Sub

Private Sub alan_hesapla()
 Dim i, j As Integer
 Dim A_(400), B_(400) As Integer
 
'ilk deðerlerini ata

 For i = 0 To 399
     A_(i) = Int(1 + ((X_(i) - bas_x) / yatay)) 'x deðerlerini kare sýralarýna(baþlangýçtan uzaklýklarý) göre ata
     B_(i) = Int(1 + ((Y_(i) - bas_y) / dusey)) 'aynýsýný y için yap
 Next
 
 
'kapladýklarýný hesapla
 kapli_alan = 400 'en fazla
 
 For i = 0 To 399
  For j = 0 To 399
   If i <> j Then
     If A_(i) > 0 And B_(i) > 0 Then
       If (A_(i) = A_(j)) And (B_(i) = B_(j)) Then
           
           kapli_alan = kapli_alan - 1
           A_(j) = 0
           B_(j) = 0
       
       End If
     End If
   End If
  Next
 Next

'buda max kapli alaný bulur
If max_kapli < kapli_alan Then
  Label11.Caption = kapli_alan
  max_kapli = kapli_alan
End If
ent_hesapla

oran1 = kapli_alan / 4
oran2 = max_kapli / 4
End Sub

Private Sub ent_hesapla()
 Dim i As Integer
 Dim p As Double
 
 lamda = 400 / kapli_alan
 hal = hal + 1
 p = lamda * Math.Exp(-lamda)
 entropi = -p * Math.Log(p)
 Label3.Caption = entropi
 
 ent_dizi(hal) = entropi
 For i = hal + 1 To 299
   ent_dizi(i) = 0
 Next
End Sub
