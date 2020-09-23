VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activate MySoftware"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   9570
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":23D2
   ScaleHeight     =   7200
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin Secureactivation.XpBs XpBs4 
      Height          =   375
      Left            =   8520
      TabIndex        =   23
      Top             =   5640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "GO BACK"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      BackColor       =   32768
      ForeColor       =   16777215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1080
      Top             =   4080
   End
   Begin Secureactivation.Xp_ProgressBar Xp_Pro 
      Height          =   255
      Left            =   2160
      TabIndex        =   21
      Top             =   5040
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      Style           =   1
   End
   Begin Secureactivation.XpBs XpBs1 
      Height          =   375
      Left            =   7800
      TabIndex        =   20
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Fill in sample"
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
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   240
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin Secureactivation.XpBs XpBs3 
      Height          =   375
      Left            =   7800
      TabIndex        =   19
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
   Begin Secureactivation.XpBs XpBs2 
      Height          =   375
      Left            =   7800
      TabIndex        =   18
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "HTML File Creater"
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DA9C42&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5400
      Picture         =   "Form3.frx":6D20
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   14
      Top             =   4200
      Width           =   495
   End
   Begin Secureactivation.XpBs command1 
      Height          =   495
      Left            =   3240
      TabIndex        =   13
      Top             =   4200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Caption         =   "Activate Product"
      ButtonStyle     =   3
      Picture         =   "Form3.frx":75EA
      PictureWidth    =   16
      PictureHeight   =   16
      PictureSize     =   0
      OriginalPicSizeW=   16
      OriginalPicSizeH=   16
      PictureHover    =   "Form3.frx":99CC
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
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   4
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   5280
      MaxLength       =   5
      TabIndex        =   3
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3120
      MaxLength       =   17
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "-------"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   22
      Top             =   4800
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form3.frx":BDAE
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   840
      TabIndex        =   17
      Top             =   6000
      Width           =   9135
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "EASY Instructions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   16
      Top             =   5640
      Width           =   3495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "In order to activate you must be connected to internet and make sure you are using genuine information"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6000
      TabIndex        =   15
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Major ID"
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
      Left            =   3600
      TabIndex        =   12
      Top             =   3840
      Width           =   855
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   3960
      X2              =   3960
      Y1              =   3600
      Y2              =   3840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   5040
      X2              =   5040
      Y1              =   3360
      Y2              =   3600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   3120
      X2              =   5040
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   3120
      X2              =   3120
      Y1              =   3360
      Y2              =   3600
   End
   Begin VB.Label Label6 
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
      Left            =   1080
      TabIndex        =   11
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address :"
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
      Left            =   800
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
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
      Left            =   1200
      TabIndex        =   9
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration ID :"
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
      Left            =   720
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form3.frx":C0C1
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
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   8295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Activate MySoftware"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   33
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Code explains how to write a Strong product ID
Private Sub command1_Click()
Dim checknow
Dim Code1 As Single
Dim i
Dim zip
Dim final
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
MsgBox ("Please Fill In All The Information!"), vbInformation, ("Registration")
Exit Sub
End If


If Len(Text1.Text) < 4 Then
    MsgBox "The Name must be more than 4 characters.", vbInformation + vbOKOnly, "Ooops"
    Exit Sub
End If

If Text5.Text = ("8546854") And Text6.Text = "64381" Then


Else
    MsgBox "Activation Failed. Please check your information", vbCritical, ("Registration")
Exit Sub
End If


For i = 1 To Len(Text1.Text) - 1
    Code1 = Format(Asc(Right(Text1.Text, Len(Text1.Text) - i)) * 2 + (79 / i) + (i + 3 / 71), "#.#")
    zip = zip & Code1
Next i
zip = Right(zip, 8)

For i = 1 To Len(zip) - 1
    Code1 = Format(Asc(Right(zip, Len(zip) - i)) * 2 + (1 / i) + (i + 1 / 9), "#00")
    final = final & Code1
Next i
final = Right(final, Len(final) - 4)
final = final & Asc(Text1)

If Text2.Text = final Then
'checks whether the file named MAJOR ID.html ( which is typed) exixts or not and
'also verifies the information contained in MAJORID.html'
    Label12.Visible = True
    Timer1.Enabled = True
    Xp_Pro.Visible = True
    'open html file with name of MJOR ID
    checknow = Inet1.OpenURL("file://" + App.Path & "/" & "files" & "/" & Text2.Text + ".html")
    Label12.Caption = "Verifying"
   Clipboard.Clear
   'copy info contained in html file
   Clipboard.SetText checknow
   Close #1
   'print the information to a ini file
   Open App.Path & "\" & "_check.ini" For Output As #1
   Print #1, Clipboard.GetText
   Close #1
   'to begin verification
   Dim regname
   Dim productid
   Dim company
   Dim email
   On Error GoTo warning
   Open App.Path & "\" & "_check.ini" For Input As #1
   Input #1, regname
   Input #1, productid
   Input #1, company
    Input #1, email
   If regname = Text1.Text And productid = Text2.Text And company = Text3.Text And email = Text4.Text Then
   MsgBox "Thank you for Activating this product.Please Restart (This is now activated because the information in HTML(majorid.html) file has matched)", vbOKOnly, "Product Activated"
   Xp_Pro.Visible = False
   Timer1.Enabled = True
   Label12.Visible = False
   Else
warning:   MsgBox "Warning- you are using illegal information and prodduct will not be activated.If you are using this information legally and you are unable to activate then please contact our customer support at support@mydomain.com for help with payment information", vbExclamation, "Warning-Illegal Product ID"
    Xp_Pro.Visible = False
   Timer1.Enabled = True
   Label12.Visible = False
   End If
    
Else
    MsgBox "Activation Failed. Please check your information", vbCritical, ("Registration")
End If
End Sub

Private Sub Timer1_Timer()
'enable progress bar
Xp_Pro.Value = Xp_Pro.Value + 1
If Xp_Pro.Value >= 100 Then
Xp_Pro.Value = 0
End If
End Sub

Private Sub XpBs1_Click()
'fill sample where html file is already in "File" folder
Text2.Text = "71031041119811983"
Text1.Text = "Sri Harish"
Text3.Text = "Microsoft"
Text4.Text = "sriharish@msn.com"
Text5.Text = "8546854"
Text6.Text = "64381"
MsgBox "HTML file with name 71031041119811983.html has already been created in files folder. This program Checks for MAJORID.html file and verifies all the information during activation."
End Sub

Private Sub XpBs2_Click()
htmlfiles.Show
End Sub

Private Sub XpBs3_Click()
Prodkeygen.Show
End Sub

Private Sub XpBs4_Click()
'go back to main form
Unload Me
Form1.Show
End Sub
