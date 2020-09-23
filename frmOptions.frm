VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   Caption         =   "Prefrences"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1920
      Width           =   1335
   End
   Begin MSComctlLib.Slider SpeedSlider 
      Height          =   555
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   979
      _Version        =   393216
      Enabled         =   0   'False
      Min             =   1
      SelStart        =   1
      Value           =   1
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CheckBox chkPause 
      Caption         =   "Pause When Filling Grid (Select Resize to see effect)"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Time (In Seconds):"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fill Speed:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fast"
      Height          =   195
      Left            =   1920
      TabIndex        =   5
      Top             =   1560
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Slow"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   345
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkPause_Click()
    
    If chkPause.Value = 1 Then
        SpeedSlider.Enabled = True
    Else
        SpeedSlider.Enabled = False
    End If

End Sub

Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    
    If chkPause.Value = 1 Then
        DrawEffect = True
        DrawSpeed = 1000 - (SpeedSlider.Value * 100)
    Else
        DrawEffect = False
    End If
    DefaultTime = Val(txtTime.Text)
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    txtTime.Text = DefaultTime
    
End Sub
