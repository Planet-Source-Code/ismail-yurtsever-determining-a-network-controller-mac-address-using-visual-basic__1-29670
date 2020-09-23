VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSerial 
      Height          =   285
      Left            =   300
      TabIndex        =   3
      Top             =   750
      Width           =   2265
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Height          =   375
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2610
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get MAC "
      Default         =   -1  'True
      Height          =   315
      Left            =   2670
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00EDA84D&
      Caption         =   "  MAC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6285
   End
   Begin VB.Image Image2 
      Height          =   60
      Left            =   0
      Picture         =   "frmMain.frx":0000
      Top             =   230
      Width           =   9750
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

        frmMac.Show 1

    

End Sub


Private Sub Command3_Click()
    End
End Sub

Private Sub Form_Load()

    Center Me

End Sub

