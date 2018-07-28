VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "MainForm"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      ItemData        =   "frmMain.frx":0000
      Left            =   5040
      List            =   "frmMain.frx":000A
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   3
      Top             =   200
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   1680
      TabIndex        =   0
      Top             =   3000
      Value           =   100
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  '����
      Caption         =   "Set BackColor"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  '����
      Caption         =   "��ũ�� ���븦 �������� ���� ������ �����غ�����!"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   2280
      Width           =   4935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Const GWL_EXSTYLE = (-20)
    Const WS_EX_LAYERED = &H80000
    Const LWA_ALPHA = &H2
    Const LWA_COLORKEY = &H2
    
Private Sub Combo1_Click()
    Dim ColorStr As String
    
    ColorStr = Combo1.Text
    
    Select Case ColorStr
        Case "Black"
            Label1.ForeColor = vbWhite
            Label2.ForeColor = vbWhite
            frmMain.BackColor = vbBlack
        Case "White"
            Label1.ForeColor = vbBlack
            Label2.ForeColor = vbBlack
            frmMain.BackColor = vbWhite
    End Select
End Sub

Private Sub HScroll1_Change()
    Dim Srclvalue As Long   '--SrclValue ������ ���� HScroll�� ���� �������ִ� �����̴�.
    Dim RGBvalue As Long    '--RGBvalue ������ ���� HScroll������ 255�� ����� �������ִ� �����̴�.
    
    Srclvalue = CInt(HScroll1.Value) 'RGBvalue���� HScroll1.Value���� ����
    RGBvalue = 255 - (Srclvalue / 255)
    
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, RGB(RGBvalue, RGBvalue, RGBvalue), RGBvalue, LWA_ALPHA Or LWA_COLORKEY
End Sub
