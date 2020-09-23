VERSION 5.00
Begin VB.Form Prodkeygen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Key Generator"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5700
   Icon            =   "Prodkeygen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1080
      TabIndex        =   7
      Top             =   1200
      Width           =   2385
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   6
      Top             =   1680
      Width           =   2205
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   4
      Top             =   1680
      Width           =   615
   End
   Begin Secureactivation.XpBs XpBs1 
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
      _extentx        =   2778
      _extenty        =   1085
      caption         =   "Generate Key"
      buttonstyle     =   3
      picture         =   "Prodkeygen.frx":000C
      font            =   "Prodkeygen.frx":08E6
      font            =   "Prodkeygen.frx":0912
      picture         =   "Prodkeygen.frx":093E
      mousepointer    =   99
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "MAJOR ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   2880
      X2              =   2880
      Y1              =   2040
      Y2              =   2880
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Produc ID :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "REG ID :"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Prodkeygen.frx":1218
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "My company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4560
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C000&
      BorderWidth     =   3
      Height          =   495
      Left            =   3840
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Height          =   375
      Left            =   3960
      Top             =   2760
      Width           =   495
   End
End
Attribute VB_Name = "Prodkeygen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub XpBs1_Click()
Dim Code1 As Single
Dim i
Dim final
Dim zip

If Len(Text1.Text) < 4 Then
    MsgBox "The Name must be more than 4 characters.", vbInformation + vbOKOnly, "Ooops"
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
Text3.Text = "8546854"
Text4.Text = "64381"
final = Right(final, Len(final) - 4)
final = final & Asc(Text1)
Text2 = final

End Sub



