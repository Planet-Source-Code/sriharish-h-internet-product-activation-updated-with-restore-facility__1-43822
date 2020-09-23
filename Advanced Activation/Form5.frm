VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product ID Restoration Program- Special Thanks to Bob for posting a comment"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9540
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form5.frx":08CA
   ScaleHeight     =   7155
   ScaleWidth      =   9540
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00DA9C42&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "Form5.frx":5218
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   32
      Top             =   1680
      Width           =   495
   End
   Begin Secureactivation.XpBs XpBs5 
      Height          =   375
      Left            =   6840
      TabIndex        =   31
      Top             =   3480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Caption         =   "Fill In sample"
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
      ForeColor       =   255
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   120
      Top             =   4800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   4320
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   960
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin Secureactivation.XpBs XpBs3 
      Height          =   375
      Left            =   6480
      TabIndex        =   29
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Copy ID"
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
      Left            =   6840
      TabIndex        =   26
      Top             =   3000
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Caption         =   """ RESTORE ""HTML File Creater"
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
      ForeColor       =   255
   End
   Begin Secureactivation.Xp_ProgressBar prog1 
      Height          =   255
      Left            =   2040
      TabIndex        =   22
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      Style           =   1
   End
   Begin Secureactivation.XpBs XpBs1 
      Height          =   495
      Left            =   2040
      TabIndex        =   21
      Top             =   4560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Caption         =   "Try To Restore"
      ButtonStyle     =   3
      Picture         =   "Form5.frx":5AE2
      PictureWidth    =   16
      PictureHeight   =   16
      PictureSize     =   0
      OriginalPicSizeW=   16
      OriginalPicSizeH=   16
      PictureHover    =   "Form5.frx":7EC4
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
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   19
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   18
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   17
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2040
      TabIndex        =   16
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2040
      TabIndex        =   14
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   13
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Top             =   2400
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00DA9C42&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      Picture         =   "Form5.frx":A2A6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   5520
      Picture         =   "Form5.frx":A4EA
      ScaleHeight     =   855
      ScaleWidth      =   975
      TabIndex        =   1
      Top             =   -120
      Width           =   975
   End
   Begin Secureactivation.XpBs XpBs4 
      Height          =   375
      Left            =   8520
      TabIndex        =   30
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
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "You must be connect to internet to restore your Product ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   720
      TabIndex        =   33
      Top             =   1800
      Width           =   5055
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "XXXX-XXXX-XXXX-XXXX-XXXX-XXX-XXX-XXX-XXX"
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
      Left            =   2040
      TabIndex        =   28
      Top             =   5760
      Width           =   4335
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   27
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Refer to Documentation.html file in the current folder"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   25
      Top             =   6720
      Width           =   3855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   720
      TabIndex        =   24
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Verifying..."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   23
      Top             =   5160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "---MM-DD-YYYY Format"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Ordering : OR Receipt Date"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Card Provider :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   270
      TabIndex        =   10
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Card No:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Email :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Company :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Reg ID :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form5.frx":CD70
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   5895
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   6480
      X2              =   9600
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   6480
      X2              =   6480
      Y1              =   840
      Y2              =   2880
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"Form5.frx":CEA3
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   6600
      TabIndex        =   3
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Privacy Statement:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Restore Product ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hi,you know many users always forget their Product ID
'and a user can restore the product id legally
'After obtaining the product id the user should
'again Activate the software, so if there is any kind
'of illegal activities,the user will be caught while activating
'because activating verifies all the information in the
'database, VOTE FOR ME!!!!
'Refer Easy Documentation for more
'°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
Private Sub XpBs1_Click()
Dim checknow

If Len(Text1.Text) > 0 And Len(Text2.Text) > 0 And Len(Text3.Text) > 0 And Len(Text4.Text) = 16 And Len(Text5.Text) > 0 And Len(Text6.Text) = 2 And Len(Text7.Text) = 2 And Len(Text8.Text) = 4 Then
'if all the infotmation is in correct format
'the credit card no.html file will be checked
'if all the information is correct in this file then
'the product id will be displayed
Timer1.Enabled = True
prog1.Visible = True ' enable progressbar
Label12.Visible = True

Timer2.Enabled = True
'connect to creditno.html file in your server

checknow = Inet1.OpenURL("file://" + App.Path & "\" & "restoredb" & "\" & Text4.Text & ".html")
Close #1
'copy the information to an ini file
Open App.Path & "\" & "restoredid.ini" For Output As #1
Clipboard.Clear
Clipboard.SetText checknow
Print #1, Clipboard.GetText
Close #1
'begin verification

Open App.Path & "\" & "restoredid.ini" For Input As #1
'the verfication will be done in following format
'use Restore HTML file creater creating html file

Dim regid, company, email, credit, provider, billdate, productid
Line Input #1, regid
If regid = Text1.Text Then
Line Input #1, company
If company = Text2.Text Then
Line Input #1, email
If email = Text3.Text Then
Line Input #1, credit
If credit = Text4.Text Then
Line Input #1, provider
If provider = Text5.Text Then
Line Input #1, billdate
'The bill date is verified
'if date typed is 30 12 2003 in each textboxes
'it will be verified as 30122003
'i suggest for you to use Restore HTML FILE CREATER
If billdate = Text6.Text + Text7.Text + Text8.Text Then
Line Input #1, productid
Label16.Caption = productid
XpBs3.Visible = True
Timer1.Enabled = False
Timer2.Enabled = False
prog1.Visible = False
Label12.Visible = False
MsgBox "Now that the product ID is restored. It must be activated again. Please activate the software now."
End If
End If
End If
End If
End If
End If
'if information is incorrect
If regid <> Text1.Text Or company <> Text2.Text Or email <> Text3.Text Or credit <> Text4.Text Or provider <> Text5.Text Or billdate <> Text6.Text + Text7.Text + Text8.Text Then
MsgBox "The information you have provided does not match or cannot be found in our server.Please check the information and then try again.", vbExclamation, "Product Id not restored"
Timer1.Enabled = False
Timer2.Enabled = False
prog1.Visible = False
Label12.Visible = False
XpBs3.Visible = False
'delete restoredid.ini
On Error Resume Next
Kill App.Path & "\" & "restoredid.ini"
Exit Sub
End If

Else
MsgBox "Invalid Information provided.Please verfiy it and try again", vbCritical, "User error"
On Error Resume Next
Kill App.Path & "\" & "restoredid.ini"
Exit Sub
End If

End Sub
Private Sub Timer1_Timer()
'progress bar program
prog1.Value = prog1.Value + 1
If prog1.Value > 99 Then
prog1.Value = 1
End If

End Sub

Private Sub Timer2_Timer()
'The connection request will be upto 20 seconds
'If the connection request is > 20 seconds
'there will be an time out error

MsgBox "Operation Timed Out.Please check your internet connection.Make sure that you have not enable and firewall profram or services.The server may also be very busy,please try after sometime.", vbCritical, "Requezst Error"
Timer1.Enabled = False 'disable progressbar
prog1.Visible = False
Label12.Visible = False
Timer2.Enabled = False 'disable timeout timer
Exit Sub
End Sub

Private Sub XpBs2_Click()
'open restore HTML file builder
Restorebuilder.Show
End Sub

Private Sub XpBs3_Click()
Clipboard.Clear
Clipboard.SetText Label16.Caption 'copy product id
MsgBox "The following Product Id has been copied--" + Label16.Caption
End Sub

Private Sub XpBs4_Click()
'go back to main form
Unload Me
Form1.Show
End Sub

Private Sub XpBs5_Click()
'Fill in sample

'****************************
'DO NOT MODIFY THIS SECTION *
'****************************
Text1.Text = "Sri Harish"
Text2.Text = "Microsoft"
Text3.Text = "sriharish@msn.com"
Text4.Text = "3333333333333333"
Text5.Text = "Visa"
Text6.Text = "30"
Text7.Text = "12"
Text8.Text = "2003"
MsgBox "HTML File: 3333333333333333.html, which is credit cardno.html, has already been created. The restore program verifies this file while restoring.", vbInformation, ":( VOTE FOR ME"
End Sub
