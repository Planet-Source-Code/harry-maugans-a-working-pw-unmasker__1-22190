VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Unmasker"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Go Unmask It!!!"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "How Do I Use This Thing!??!?!?!"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "5"
      Top             =   240
      Width           =   855
   End
   Begin VB.Timer tmrHack 
      Enabled         =   0   'False
      Left            =   600
      Top             =   1680
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2520
      Shape           =   2  'Oval
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   10
      X1              =   3840
      X2              =   3840
      Y1              =   480
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   10
      X1              =   3120
      X2              =   3120
      Y1              =   480
      Y2              =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Timer Interval:"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim process         As Boolean

Private Declare Function GetCursorPos Lib "user32" _
(lpPoint As POINTAPI) As Long

Private Declare Function WindowFromPoint Lib "user32" _
(ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias _
"GetClassNameA" (ByVal hwnd As Long, ByVal _
lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Private Type POINTAPI
X As Long
Y As Long
End Type

Private Const EM_SETPASSWORDCHAR = &HCC

Private Sub ShowPassword()

Dim curWindow   As Long
Dim sClassName  As String * 255
Dim sName       As String
Dim lpPoint     As POINTAPI

Call GetCursorPos(lpPoint)
curWindow = WindowFromPoint(lpPoint.X, lpPoint.Y)

Call GetClassName(curWindow, sClassName, 255)
sName = Trim(Left(sClassName, InStr(sClassName, vbNullChar) - 1))

If sName = "Edit" Or InStr(sName, "TextBox") > 0 Then
Call SendMessage(curWindow, EM_SETPASSWORDCHAR, 0, 0)
Me.SetFocus
MsgBox "Hacked!", vbCritical, "Yay!"
Else
MsgBox "Sorry! Cannot get the password", vbExclamation, "Error !"
End If

tmrHack.Enabled = False

End Sub
    
    
Private Sub Command1_Click()
    MsgBox "Type in the textbox the number of seconds that you want to have before it activates.  Click on the button 'Go Hack It' and then you have the number of seconds in the text box to goto an application with a password masked textbox (like with astericks).  Click in the masked textbox and hold your mouse over it until the time runs out.  It will then tell you the outcome of the hack attempt.  It all goes correctly, when you click back on this program, then the hacked application again, the textbox will be unmasked.  Enjoy! =)              (Programmed by: Harry Maugans  while testing advanced API calls.) (This program is only for use in Switzerland and Iraq, neither of which support US copyright or trademark laws.) (Copyleft 2001, Messed Up Studios, Inc.)", vbInformation, "Help"
End Sub

Private Sub Command2_Click()
    tmrHack.Interval = Val(Text1.Text) * 1000
    tmrHack.Enabled = True
End Sub

Private Sub tmrHack_Timer()
    Call ShowPassword
End Sub
