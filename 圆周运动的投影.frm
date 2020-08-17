VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "圆周运动的投影CAI课件"
   ClientHeight    =   4860
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   7896
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   7896
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2604
      Top             =   4044
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      Left            =   216
      Max             =   100
      TabIndex        =   1
      Top             =   4404
      Width           =   2352
   End
   Begin VB.PictureBox Tu 
      Height          =   3864
      Left            =   3228
      ScaleHeight     =   3816
      ScaleWidth      =   4560
      TabIndex        =   0
      Top             =   132
      Width           =   4608
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   48
         X2              =   4848
         Y1              =   1920
         Y2              =   1920
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "结论：圆周运动在圆平面方向上的投影是简谐振动！"
      Height          =   180
      Left            =   3336
      TabIndex        =   3
      Top             =   4452
      Width           =   4140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Left            =   1092
      TabIndex        =   2
      Top             =   4176
      Width           =   576
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   456
      X2              =   2724
      Y1              =   2124
      Y2              =   2124
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   204
      Left            =   36
      Shape           =   3  'Circle
      Top             =   1992
      Width           =   204
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim delta As Single '圆周运动的投影
Private Sub Form_Load()
  Tu.AutoRedraw = True
  Tu.BackColor = vbBlack
  Line1.Y1 = Tu.Height / 2
  Line1.Y2 = Tu.Height / 2
  HScroll1.Value = 10
  delta = 3.1416 / 360 * HScroll1.Value
End Sub
Private Sub HScroll1_Change()
 delta = 3.1416 / 360 * HScroll1.Value
 Label1.Caption = "速度值=" & HScroll1.Value
End Sub
Private Sub Timer1_Timer()
  Static alfa As Single
  Dim y As Integer, oldy As Integer, x As Single
  Const dx = 20
  oldy = Tu.Height / 3 * Sin(alfa) + Tu.Height / 2
  alfa = alfa + delta
  y = Tu.Height / 3 * Sin(alfa) + Tu.Height / 2
  Tu.Line (dx, oldy)-(0, y), vbGreen
  Tu.Picture = Tu.Image
  Tu.Line (dx, oldy)-(0, y), Tu.BackColor
  Tu.PaintPicture Tu.Picture, dx, 0
  x = Tu.Height / 3 * Cos(alfa) + Tu.Left / 2
  Form1.PSet (x + Shape1.Width / 2, y + Shape1.Width)
  Shape1.Left = x
  Shape1.Top = y + Tu.Top
  Line2.X1 = x + Shape1.Width
  Line2.X2 = Tu.Left + Tu.Top
  Line2.Y1 = y + Shape1.Width
  Line2.Y2 = y + Shape1.Width
End Sub

