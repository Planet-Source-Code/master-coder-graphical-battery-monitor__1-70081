VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frmVisibility 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Battery Monitor - Desktop Battery Visibility Level"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3690
      TabIndex        =   0
      Top             =   2175
      Width           =   915
   End
   Begin ComctlLib.Slider sldVisibility 
      Height          =   555
      Left            =   180
      TabIndex        =   1
      Top             =   210
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   979
      _Version        =   327682
      LargeChange     =   10
      SmallChange     =   10
      Min             =   10
      Max             =   100
      SelStart        =   10
      TickFrequency   =   10
      Value           =   10
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Level:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   1470
      TabIndex        =   8
      Top             =   1845
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visibility Level:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   1470
      TabIndex        =   7
      Top             =   1575
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   1365
      TabIndex        =   6
      Top             =   915
      Width           =   1620
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   1365
      X2              =   4590
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Style:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   1470
      TabIndex        =   5
      Top             =   1305
      Width           =   465
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Style1 (Green)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   2685
      TabIndex        =   4
      Top             =   1305
      Width           =   1320
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2775
      TabIndex        =   3
      Top             =   1575
      Width           =   60
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   2775
      TabIndex        =   2
      Top             =   1845
      Width           =   60
   End
End
Attribute VB_Name = "frmVisibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oDIB As New c32bppDIB

Private Sub cmdOK_Click()

    Unload Me

End Sub

Private Sub DrawSample()

    oDIB.Render Me.hdc, -20, 50, , , , , , , sldVisibility.Value

End Sub

Private Sub Form_Load()

    sldVisibility.Value = ((bTransparency / 25) * 10)

    lblInfo(1).Caption = sldVisibility.Value

    lblInfo(2).Caption = sldVisibility.Value

    oDIB.LoadPicture_File (App.Path & "\style1\Battery-100.png")

    DrawSample

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set oDIB = Nothing

    bTransparency = ((sldVisibility.Value / 10) * 25)

End Sub

Private Sub sldVisibility_Scroll()

    Me.Cls

    lblInfo(2).Caption = sldVisibility.Value

    DrawSample

End Sub

