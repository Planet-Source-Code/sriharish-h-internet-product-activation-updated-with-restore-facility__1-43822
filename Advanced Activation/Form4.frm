VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   6975
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "Form4.frx":08CA
      Top             =   1320
      Width           =   5055
   End
   Begin Secureactivation.XpBs XpBs1 
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "Email Me"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      URL             =   "mailto:sriharish@msn.com"
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Special Thanks to : Michael Jackson, Bob, E.Murphy, Harry guru"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   4575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Original Author : Myself"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "This is an Updated Version :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "VOTE FOR ME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   2055
      Left            =   5280
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      X1              =   0
      X2              =   6840
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"Form4.frx":0DC3
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "---"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "To :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub Form_Load()
Dim username As String * 30
Dim returns
returns = GetUserName(username, 30)
username = Left(username, InStr(username, Chr(0)) - 1)
Label2.Caption = username
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Dim abc
 Dim efg
 
 If abc = MsgBox("Did you vote?", vbQuestion + vbYesNo, ":(") = vbYes Then
Form1.Show
Unload Me

End If
 
 
End Sub
