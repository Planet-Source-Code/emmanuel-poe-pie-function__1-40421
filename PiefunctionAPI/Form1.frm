VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Pie Demonstration"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   409
   ScaleMode       =   0  'User
   ScaleWidth      =   394
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Draw both"
      Height          =   330
      Left            =   2160
      TabIndex        =   41
      Top             =   4050
      Width           =   1680
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   2250
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   40
      Top             =   4545
      Width           =   1515
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Draw"
      Height          =   330
      Left            =   3990
      TabIndex        =   27
      Top             =   4995
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   1530
      Left            =   4095
      ScaleHeight     =   100
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   26
      Top             =   3060
      Width           =   1530
   End
   Begin VB.TextBox Text16 
      Enabled         =   0   'False
      Height          =   330
      Left            =   5250
      TabIndex        =   21
      Text            =   "0"
      Top             =   1980
      Width           =   540
   End
   Begin VB.TextBox Text15 
      Enabled         =   0   'False
      Height          =   330
      Left            =   4515
      TabIndex        =   20
      Text            =   "0"
      Top             =   1980
      Width           =   540
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      Height          =   330
      Left            =   3780
      TabIndex        =   19
      Text            =   "100"
      Top             =   1980
      Width           =   540
   End
   Begin VB.TextBox Text13 
      Enabled         =   0   'False
      Height          =   330
      Left            =   3045
      TabIndex        =   18
      Text            =   "50"
      Top             =   1980
      Width           =   540
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   330
      Left            =   2310
      TabIndex        =   17
      Text            =   "100"
      Top             =   1980
      Width           =   540
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   330
      Left            =   1575
      TabIndex        =   16
      Text            =   "100"
      Top             =   1980
      Width           =   540
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   330
      Left            =   840
      TabIndex        =   15
      Text            =   "0"
      Top             =   1980
      Width           =   540
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   330
      Left            =   105
      TabIndex        =   14
      Text            =   "0"
      Top             =   1980
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw"
      Height          =   330
      Left            =   270
      TabIndex        =   9
      Top             =   4995
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Height          =   330
      Left            =   5250
      TabIndex        =   8
      Text            =   "100"
      Top             =   735
      Width           =   540
   End
   Begin VB.TextBox Text7 
      Height          =   330
      Left            =   4515
      TabIndex        =   7
      Text            =   "50"
      Top             =   735
      Width           =   540
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   330
      Left            =   3780
      TabIndex        =   6
      Text            =   "0"
      Top             =   735
      Width           =   540
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   330
      Left            =   3045
      TabIndex        =   5
      Text            =   "0"
      Top             =   735
      Width           =   540
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   330
      Left            =   2310
      TabIndex        =   4
      Text            =   "100"
      Top             =   735
      Width           =   540
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   330
      Left            =   1575
      TabIndex        =   3
      Text            =   "100"
      Top             =   735
      Width           =   540
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   330
      Left            =   810
      TabIndex        =   2
      Text            =   "0"
      Top             =   720
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Text            =   "0"
      Top             =   735
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   1515
      Left            =   360
      ScaleHeight     =   100
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   3060
      Width           =   1515
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Mouse Point"
      Height          =   240
      Left            =   2475
      TabIndex        =   39
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label19 
      Height          =   555
      Left            =   2520
      TabIndex        =   38
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Label Label18 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5715
      TabIndex        =   37
      Top             =   3555
      Width           =   195
   End
   Begin VB.Label Label17 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   36
      Top             =   3510
      Width           =   195
   End
   Begin VB.Label Label16 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4725
      TabIndex        =   35
      Top             =   4635
      Width           =   195
   End
   Begin VB.Label Label15 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   945
      TabIndex        =   34
      Top             =   4635
      Width           =   195
   End
   Begin VB.Label Label14 
      Caption         =   "   X              Y               X              Y             X              Y               X              Y"
      Height          =   240
      Left            =   90
      TabIndex        =   33
      Top             =   2340
      Width           =   5730
   End
   Begin VB.Label Label13 
      Caption         =   "   X              Y               X              Y             X              Y               X              Y"
      Height          =   240
      Left            =   135
      TabIndex        =   32
      Top             =   1080
      Width           =   5685
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   390
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Pie 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3990
      TabIndex        =   31
      Top             =   2730
      Width           =   1695
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Pie 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   285
      TabIndex        =   30
      Top             =   2730
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Pie 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   29
      Top             =   1350
      Width           =   5685
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Pie 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   28
      Top             =   105
      Width           =   5685
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Upper Left"
      Height          =   225
      Left            =   105
      TabIndex        =   25
      Top             =   1665
      Width           =   1275
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Lower Right"
      Height          =   225
      Left            =   1575
      TabIndex        =   24
      Top             =   1665
      Width           =   1275
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "StartPoint Radial1"
      Height          =   225
      Left            =   3045
      TabIndex        =   23
      Top             =   1665
      Width           =   1275
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "EndPoint Radial2"
      Height          =   225
      Left            =   4515
      TabIndex        =   22
      Top             =   1665
      Width           =   1275
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "EndPoint Radial2"
      Height          =   225
      Left            =   4515
      TabIndex        =   13
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "StartPoint Radial1"
      Height          =   225
      Left            =   3045
      TabIndex        =   12
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Lower Right"
      Height          =   225
      Left            =   1575
      TabIndex        =   11
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Upper Left"
      Height          =   225
      Left            =   105
      TabIndex        =   10
      Top             =   420
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long, ByVal x3 As Long, ByVal Y3 As Long, ByVal x4 As Long, ByVal Y4 As Long) As Long

Private Sub Command1_Click()
Dim x1, x2, x3, x4, x5, x6, x7, x8 As Long
Picture1.Cls
Picture1.AutoRedraw = True
Picture1.DrawMode = vbCopyPen
Picture1.DrawStyle = vbSolid
Picture1.DrawWidth = 2
Picture1.FillColor = vbRed
Picture1.FillStyle = vbSolid
Picture1.ForeColor = vbBlack
    
    x1 = Val(Text1.Text)
    x2 = Val(Text2.Text)
    x3 = Val(Text3.Text)
    x4 = Val(Text4.Text)
    x5 = Val(Text5.Text)
    x6 = Val(Text6.Text)
    x7 = Val(Text7.Text)
    x8 = Val(Text8.Text)
    Picture1.FillColor = vbRed
    Pie Picture1.hdc, x1, x2, x3, x4, x5, x6, x7, x8
End Sub

Private Sub Command2_Click()
Dim x9, x10, x11, x12, x13, x14, x15, x16 As Long
Picture2.Cls
Picture2.AutoRedraw = True
Picture2.DrawMode = vbCopyPen
Picture2.DrawStyle = vbSolid
Picture2.DrawWidth = 2
Picture2.FillColor = vbBlue
Picture2.FillStyle = vbSolid
Picture2.ForeColor = vbBlack
    x9 = Val(Text9.Text)
    x10 = Val(Text10.Text)
    x11 = Val(Text11.Text)
    x12 = Val(Text12.Text)
    x13 = Val(Text13.Text)
    x14 = Val(Text14.Text)
    x15 = Val(Text15.Text)
    x16 = Val(Text16.Text)
    Picture2.FillColor = vbBlue
    Pie Picture2.hdc, x9, x10, x11, x12, x13, x14, x15, x16
End Sub

Private Sub Command3_Click()
Dim x7, x8 As Long
Picture3.Cls
Picture3.AutoRedraw = True
Picture3.DrawMode = vbCopyPen
Picture3.DrawStyle = vbSolid
Picture3.DrawWidth = 2
Picture3.FillColor = vbRed
Picture3.FillStyle = vbSolid
Picture3.ForeColor = vbBlack
    
    x7 = Val(Text7.Text)
    x8 = Val(Text8.Text)
    Picture3.FillColor = vbRed
    Pie Picture3.hdc, 0, 0, 100, 100, 0, 0, x7, x8
    ''''''''''''''''''''''''''''''''''''''''''''''''
    Picture3.FillColor = vbBlue
    Pie Picture3.hdc, 0, 0, 100, 100, x7, x8, 0, 0
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.Caption = ""
Label19.Caption = "x= " & X & " Y= " & Y
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.Caption = ""
Label19.Caption = "x= " & X & " Y= " & Y
End Sub

Private Sub Text7_Change()
Text13.Text = Text7.Text
End Sub

Private Sub Text8_Change()
Text14.Text = Text8.Text
End Sub
