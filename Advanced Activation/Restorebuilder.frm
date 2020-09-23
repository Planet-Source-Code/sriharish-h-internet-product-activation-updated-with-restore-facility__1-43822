VERSION 5.00
Begin VB.Form Restorebuilder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restore HTML FILE BUILDER"
   ClientHeight    =   5580
   ClientLeft      =   2760
   ClientTop       =   3855
   ClientWidth     =   5130
   Icon            =   "Restorebuilder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Restorebuilder.frx":1CFA
   ScaleHeight     =   5580
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   2520
      TabIndex        =   21
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2040
      TabIndex        =   20
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox pid 
      Height          =   285
      Left            =   1560
      TabIndex        =   19
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1560
      TabIndex        =   18
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1560
      TabIndex        =   17
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1560
      MaxLength       =   16
      TabIndex        =   16
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   15
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   14
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   960
      Width           =   2655
   End
   Begin Secureactivation.XpBs XpBs2 
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   4440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Caption         =   "Key Generator"
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
   End
   Begin Secureactivation.XpBs XpBs1 
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   4440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Caption         =   "Build HTML File"
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DA9C42&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   4815
      Begin VB.Label Label9 
         BackColor       =   &H00DA9C42&
         Caption         =   "Product ID Example :"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00DA9C42&
         Caption         =   "85264-3530964363463634-64381"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Note the Credit card no should be 16 digits. Generate REG ID and Product Id using Key generator."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   5040
      Width           =   4935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Restore HTML BUILDER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   720
      TabIndex        =   10
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Card Provider :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Card No :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Email :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Company :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "REG ID :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "Restorebuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub XpBs1_Click()
If Len(Text1.Text) > 0 And Len(Text2.Text) > 0 And Len(Text3.Text) > 0 And Len(Text4.Text) = 16 And Len(Text5.Text) > 0 And Len(Text6.Text) = 2 And Len(Text7.Text) = 2 And Len(Text8.Text) = 4 And Len(pid.Text) > 0 Then
Close #1
'create creditcardno.html file
Open App.Path & "\" & "restoredb" & "\" & Text4.Text + ".html" For Output As #1
Print #1, Text1.Text
Print #1, Text2.Text
Print #1, Text3.Text
Print #1, Text4.Text
Print #1, Text5.Text
'print bill date
Print #1, Text6.Text + Text7.Text + Text8.Text
Print #1, pid.Text
Close #1
MsgBox "HTML file: " + Text4.Text + ".html" + " has been created in restoredb folder. The Program verifies for this file while restoring.", vbInformation, ":)VOTE FOR ME"
Else
MsgBox "You have provided invalid information. Make ure credit card no is 16 digits.", vbExclamation, ":)VOTE FOR ME"
Exit Sub
End If

End Sub

Private Sub XpBs2_Click()
'show key generator
Prodkeygen.Show
End Sub
