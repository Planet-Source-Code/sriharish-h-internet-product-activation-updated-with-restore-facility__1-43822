VERSION 5.00
Begin VB.Form htmlfiles 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "*HTML File Creater*"
   ClientHeight    =   3060
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5535
   Icon            =   "htmlfiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Secureactivation.XpBs XpBs1 
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   2520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Caption         =   "Create HTML FIle"
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Email :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Company :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MAJOR Product ID :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "REG ID :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"htmlfiles.frx":000C
      ForeColor       =   &H00004080&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "htmlfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub XpBs1_Click()
'this will create the html file with name of MAJOR PROD ID in files folder
Close #1
If Len(Text1.Text) > 0 And Len(Text2.Text) > 0 And Len(Text3.Text) > 0 And Len(Text4.Text) > 0 Then
Open App.Path & "\" & "files" & "\" & Text2.Text + ".html" For Output As #1
Print #1, Text1.Text
Print #1, Text2.Text
Print #1, Text3.Text
Print #1, Text4.Text
MsgBox "HTML file Name :" + Text2.Text + ".html" + " has been created", vbInformation, "HTML FILE CREATER"
Else
MsgBox "Fill in all the information fist", vbInformation, "HTML File Creater"
End If

End Sub
