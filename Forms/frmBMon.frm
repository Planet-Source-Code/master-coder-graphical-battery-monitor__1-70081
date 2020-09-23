VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBMon 
   Appearance      =   0  'Flat
   BackColor       =   &H0080C0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   2760
   ClientLeft      =   -30
   ClientTop       =   315
   ClientWidth     =   3135
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "frmBMon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSComctlLib.ImageList IL 
      Left            =   1710
      Top             =   5985
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   255
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   107
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":02A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":03FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":0558
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":06B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":080C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":0966
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":0AC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":0C1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":0D74
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":0ECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":1028
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":1182
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":12DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":1436
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":1590
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":16EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":1844
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":199E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":1AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":1C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":1DAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":1F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":2060
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":21BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":2314
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":246E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":25C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":2722
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":287C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":29D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":2B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":2C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":2DE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":2F3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":3098
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":31F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":334C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":34A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":3600
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":375A
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":38B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":3A0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":3B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":3CC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":3E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":3F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":40D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":422A
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":4384
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":44DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":4638
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":4792
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":48EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":4A46
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":4BA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":4CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":4E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":4FAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":5108
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":5262
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":53BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":5516
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":5670
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":57CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":5924
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":5A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":5BD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":5D32
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":5E8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":5FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":6140
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":629A
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":63F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":654E
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":66A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":6802
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":695C
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":6AB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":6C10
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":6D6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":6EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":701E
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":7178
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":72D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":742C
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":7586
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":76E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":783A
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":7994
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":7AEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":7C48
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":7DA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":7EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":8056
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":81B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":830A
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":8464
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":85BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":8718
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":8872
            Key             =   "Ch1"
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":89CC
            Key             =   "Ch2"
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":8B26
            Key             =   "Ch3"
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":8C80
            Key             =   "Ch4"
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":8DDA
            Key             =   "Blank"
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":8F34
            Key             =   "AC"
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMon.frx":908E
            Key             =   "Bolt"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1230
      Top             =   5985
   End
   Begin VB.Menu mnuPop 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuDontShowBattery 
         Caption         =   "Hide Desktop Battery Icon"
      End
      Begin VB.Menu mnuOnTop 
         Caption         =   "Set Battery as Top Most Window"
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTimer 
         Caption         =   "Timer Interval"
         Index           =   0
         Begin VB.Menu mnuTimeInterval 
            Caption         =   "---- Step Animation ----"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuTimeInterval 
            Caption         =   "750 ms"
            Index           =   1
         End
         Begin VB.Menu mnuTimeInterval 
            Caption         =   "---- Lightening Bolt ----"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu mnuTimeInterval 
            Caption         =   "750 ms"
            Index           =   3
         End
         Begin VB.Menu mnuTimeInterval 
            Caption         =   "2 secs"
            Index           =   4
         End
         Begin VB.Menu mnuTimeInterval 
            Caption         =   "---- No Animation ----"
            Enabled         =   0   'False
            Index           =   5
         End
         Begin VB.Menu mnuTimeInterval 
            Caption         =   "30 secs"
            Index           =   6
         End
         Begin VB.Menu mnuTimeInterval 
            Caption         =   "60 secs"
            Index           =   7
         End
      End
      Begin VB.Menu mnuSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIcon 
         Caption         =   "Desktop Icon Size"
         Begin VB.Menu mnuIconSize 
            Caption         =   "Custom"
            Index           =   0
         End
         Begin VB.Menu mnuIconSize 
            Caption         =   "256 x 256"
            Index           =   1
         End
         Begin VB.Menu mnuIconSize 
            Caption         =   "128 x 128"
            Index           =   2
         End
         Begin VB.Menu mnuIconSize 
            Caption         =   "96 x96"
            Index           =   3
         End
         Begin VB.Menu mnuIconSize 
            Caption         =   "64 x 64"
            Index           =   4
         End
         Begin VB.Menu mnuIconSize 
            Caption         =   "48 x 48"
            Index           =   5
         End
         Begin VB.Menu mnuIconSize 
            Caption         =   "32 x 32"
            Index           =   6
         End
      End
      Begin VB.Menu mnuTransparent 
         Caption         =   "Visibility"
         Begin VB.Menu mnuShowSelectionBox 
            Caption         =   "Show Visibility Wizard"
         End
         Begin VB.Menu mnuSep10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTrans 
            Caption         =   "10%"
            Index           =   1
         End
         Begin VB.Menu mnuTrans 
            Caption         =   "20%"
            Index           =   2
         End
         Begin VB.Menu mnuTrans 
            Caption         =   "30%"
            Index           =   3
         End
         Begin VB.Menu mnuTrans 
            Caption         =   "40%"
            Index           =   4
         End
         Begin VB.Menu mnuTrans 
            Caption         =   "50%"
            Index           =   5
         End
         Begin VB.Menu mnuTrans 
            Caption         =   "60%"
            Index           =   6
         End
         Begin VB.Menu mnuTrans 
            Caption         =   "70%"
            Index           =   7
         End
         Begin VB.Menu mnuTrans 
            Caption         =   "80%"
            Index           =   8
         End
         Begin VB.Menu mnuTrans 
            Caption         =   "90%"
            Index           =   9
         End
         Begin VB.Menu mnuTrans 
            Caption         =   "100%"
            Index           =   10
         End
      End
      Begin VB.Menu mnuSep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRun 
         Caption         =   "Start with Windows"
      End
      Begin VB.Menu mnuRunChecks 
         Caption         =   "Don't perform Start-up Checks"
      End
      Begin VB.Menu mnuSep07 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuSep05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnload 
         Caption         =   "Unload"
      End
   End
End
Attribute VB_Name = "frmBMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------------------------------------------------
' PROJECT    : Battery Monitor
'
' FILENAME   : Battery.exe
' CREATED BY : Bryan Utley
'         ON : Tuesday, January 8, 2008 at 10:15:18 AM
'
' COPYRIGHT  : Copyright 2008 - All Rights Reserved
'              The World Wide Web Programmer's Consortium
'
' DESCRIPTION: Provide a graphical display of the battery level for laptop
'              computers.
'
' COMMENTS   : N/A
'
' WEB SITE   : http://www.thewwwpc.com
' E-MAIL     : bryan@thewwwpc.com
'
' VERSIONS   :                          Freeware     Professional
'                                       -------------------------
'              Multiple Battery Styles      -             X
'              Log battery usage            -             X
'              Track Multiple Batteries     -             X
'              Automatic Save               -             X
'              Automatic Updates            -             X
'
' MODIFICATION HISTORY:
'
' 1.0.0   MODIFIED ON   : Tuesday, February 5, 2008 at 11:01:48 PM
'         MODIFIED BY   : Bryan Utley
'         MODIFICATIONS : Initial Version
'
' ---------------------------------------------------------------------------------------------------------------------------------
'
Option Explicit

Private Const sAppName  As String = "Battery Monitor"

Private sTooltipInfo    As String
Private DesktopBattery  As New frmDesktopDisplay

Private Sub ClearTimes()

    mnuTimeInterval(1).Checked = False
    mnuTimeInterval(3).Checked = False
    mnuTimeInterval(4).Checked = False
    mnuTimeInterval(6).Checked = False
    mnuTimeInterval(7).Checked = False

End Sub

Private Sub ClearTransparency()

  Dim i As Integer

    For i = 1 To 10

        mnuTrans(i).Checked = False

    Next

End Sub

Private Sub Form_Load()

   Dim i As Integer
   
    If App.PrevInstance = True Then End

    frmSplash.Show

    iPowerLevel = 100
    bAttachedToAC = True

    GetSystemPowerStatus SysPower

    With nID
        .cbSize = Len(nID)
        .hwnd = Me.hwnd
        .uid = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = IL.ListImages("Blank").Picture
    End With

    UpdateTrayToolTip

    If Shell_NotifyIcon(NIM_ADD, nID) = 1 Then
        Tray = True
    End If

    LoadOptions
    
    LoadStartUpOptions

    With DesktopBattery

        If bShowDesktopIcon Then
            .Show
        End If

        .DesktopBatteryTopMost = bDesktopBatteryTopMost
        
        .UpdateImageSize 0, iIconSize

        If iIconSize = 256 Then
            .iImageSizeIndex = 1
         ElseIf iIconSize = 128 Then
            .iImageSizeIndex = 2
         ElseIf iIconSize = 96 Then
            .iImageSizeIndex = 3
         ElseIf iIconSize = 72 Then
            .iImageSizeIndex = 4
         ElseIf iIconSize = 48 Then
            .iImageSizeIndex = 5
         ElseIf iIconSize = 32 Then
            .iImageSizeIndex = 6
         Else
            .iImageSizeIndex = 0
        End If

        For i = mnuIconSize.LBound To mnuIconSize.UBound
            mnuIconSize(i).Checked = IIf(i = .iImageSizeIndex, True, False)
        Next

    End With

    If Timer.Interval = 0 Then
        Timer.Interval = 750
    End If

    Timer.Enabled = True
    Call Timer_Timer

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim Msg   As Long

    Msg = IIf(Me.ScaleMode = vbPixels, X, X / Screen.TwipsPerPixelX)

    Select Case Msg
     Case WM_RBUTTONUP
        
        PopupMenu mnuPop, , , , mnuUnload

     Case WM_LBUTTONDBLCLK

        If MsgBox("Do you wish to close Battery Monitor?", vbQuestion Or vbYesNo, "Exit Battery Monitor...") = vbYes Then
            Unload Me
        End If

    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Dim frm As Form

    If Tray Then
        Shell_NotifyIcon NIM_DELETE, nID
    End If

    UpdateKey HKEY_LOCAL_MACHINE, RegKey, "Top", CStr(DesktopBattery.Top)
    UpdateKey HKEY_LOCAL_MACHINE, RegKey, "Left", CStr(DesktopBattery.Left)
    UpdateKey HKEY_LOCAL_MACHINE, RegKey, "Size", CStr(iIconSize)
    UpdateKey HKEY_LOCAL_MACHINE, RegKey, "Transparency", CStr(bTransparency)

    For Each frm In Forms
        Unload frm
    Next frm

End Sub

Private Function GetFile2Run()

  Dim FileName As String

    FileName = App.Path & "\" & App.EXEName & ".exe"
    Validate FileName
    GetFile2Run = FileName

End Function

Private Sub LoadOptions()

    bTransparency = Val(GetKeyValue(HKEY_LOCAL_MACHINE, RegKey, "Transparency"))

    If bTransparency < 10 Then
        bTransparency = 100
    End If

    mnuTrans(bTransparency / 25).Checked = True

    Select Case Val(GetKeyValue(HKEY_LOCAL_MACHINE, RegKey, "Animate"))
     
     Case 0
        Animate = 0

     Case 1
        Animate = 1

     Case Else
        Animate = 2

    End Select

    Select Case Val(GetKeyValue(HKEY_LOCAL_MACHINE, RegKey, "Interval"))
     
     Case 750
        If Animate = 1 Then
            mnuTimeInterval(3).Checked = True
         Else
            mnuTimeInterval(1).Checked = True
        End If

        Timer.Interval = 750

     Case 2000
        mnuTimeInterval(4).Checked = True
        Timer.Interval = 2000

     Case 30000
        mnuTimeInterval(6).Checked = True
        Timer.Interval = 30000

     Case 60000
        mnuTimeInterval(7).Checked = True
        Timer.Interval = 60000

     Case Else
        Timer.Interval = 750
        Animate = 2
        mnuTimeInterval(1).Checked = True

    End Select

End Sub

Private Sub LoadStartUpOptions()

  Dim File2Run As String
  Dim mBoxRslt As VbMsgBoxResult

    On Error Resume Next

    File2Run = GetKeyValue(HKEY_LOCAL_MACHINE, RunKey, "BBMon")

    With DesktopBattery
        
        .Move CSng(GetKeyValue(HKEY_LOCAL_MACHINE, RegKey, "Left")), CSng(GetKeyValue(HKEY_LOCAL_MACHINE, RegKey, "Top"))
        
        iIconSize = CSng(GetKeyValue(HKEY_LOCAL_MACHINE, RegKey, "Size"))
        
        mnuOnTop.Checked = IIf(GetKeyValue(HKEY_LOCAL_MACHINE, RegKey, "OnTop") = "1", True, False)
        
        .DesktopBatteryTopMost = mnuOnTop.Checked

        If GetKeyValue(HKEY_LOCAL_MACHINE, RegKey, "ShowIcon") = "" Or GetKeyValue(HKEY_LOCAL_MACHINE, RegKey, "ShowIcon") = "0" Then
            bShowDesktopIcon = True
         Else
            bShowDesktopIcon = False
        End If

        mnuDontShowBattery.Checked = Not bShowDesktopIcon

    End With

    If Val(GetKeyValue(HKEY_LOCAL_MACHINE, RegKey, "NoPrompt")) <> 0 Then
        mnuRunChecks.Checked = True
    End If

    If (File2Run <> "" And Dir(File2Run) = "") Or Dir(File2Run) = "" Or File2Run = "" Then
        
        If mnuRunChecks.Checked = False Then
            mBoxRslt = MsgBox("This Battery Monitor is not set to automatically start when windows starts.  " & _
                              "You can have the Battery Monitor start up automatically. To set this " & _
                              "up, press 'YES'. Otherwise, press 'NO'. If you prefer not to be asked this " & _
                              "question again, press 'CANCEL'", vbQuestion Or vbYesNoCancel, "Auto Run when Windows Starts?")

            Select Case mBoxRslt
             Case vbYes
                File2Run = GetFile2Run
                UpdateKey HKEY_LOCAL_MACHINE, RunKey, "BBMon", File2Run
                mnuRun.Checked = True

             Case vbCancel
                UpdateKey HKEY_LOCAL_MACHINE, RegKey, "NoPrompt", "1"
                mnuRunChecks.Checked = True

            End Select

        End If
     
     Else
        mnuRun.Checked = True

    End If

End Sub

Private Sub mnuAbout_Click()

    frmSplash.Show

End Sub

Private Sub mnuDontShowBattery_Click()

    mnuDontShowBattery.Checked = Not mnuDontShowBattery.Checked
    
    UpdateKey HKEY_LOCAL_MACHINE, RegKey, "ShowIcon", IIf(mnuDontShowBattery.Checked = True, "1", "0")

    If mnuDontShowBattery.Checked = True Then
        DesktopBattery.Hide
     Else
        DesktopBattery.Show
    End If

End Sub

Private Sub mnuIconSize_Click(Index As Integer)

  Dim i As Integer

    For i = mnuIconSize.LBound To mnuIconSize.UBound
        mnuIconSize(i).Checked = IIf(i = Index, True, False)
    Next

    If Index = 0 Then
        DesktopBattery.ShowCustomSizeWindow
     Else
        DesktopBattery.UpdateImageSize Index
    End If

End Sub

Private Sub mnuOnTop_click()

    mnuOnTop.Checked = Not mnuOnTop.Checked
    
    UpdateKey HKEY_LOCAL_MACHINE, RegKey, "OnTop", IIf(mnuOnTop.Checked = True, "1", "0")
    
    DesktopBattery.DesktopBatteryTopMost = mnuOnTop.Checked

End Sub

Private Sub mnuRunChecks_Click()

    mnuRunChecks.Checked = Not mnuRunChecks.Checked
    
    UpdateKey HKEY_LOCAL_MACHINE, RegKey, "NoPrompt", IIf(mnuRunChecks.Checked = True, "1", "0")

End Sub

Private Sub mnuRun_Click()

    mnuRun.Checked = Not mnuRun.Checked
    
    UpdateKey HKEY_LOCAL_MACHINE, RunKey, "BBMon", IIf(mnuRun.Checked = True, GetFile2Run, vbNullString)

End Sub

Private Sub mnuShowSelectionBox_Click()

    frmVisibility.Show vbModal, Me

    ClearTransparency

    DesktopBattery.PaintImage

End Sub

Private Sub mnuTimeInterval_Click(Index As Integer)

    Call ClearTimes

    Select Case Index

     Case 1
        Timer.Interval = 750
        Animate = 2

     Case 3
        Timer.Interval = 750
        Animate = 1

     Case 4
        Timer.Interval = 2000
        Animate = 1

     Case 6
        Timer.Interval = 30000
        Animate = 0

     Case 7
        Timer.Interval = 60000
        Animate = 0

    End Select

    mnuTimeInterval(Index).Checked = True

    Call UpdateInterval

End Sub

Private Sub mnuTrans_Click(Index As Integer)

    ClearTransparency

    mnuTrans(Index).Checked = True

    bTransparency = Index * 25

    DesktopBattery.PaintImage

End Sub

Private Sub mnuUnload_Click()

    Unload Me

End Sub

Private Sub Timer_Timer()

  Static Y As Integer

  Dim X    As Integer

    GetSystemPowerStatus SysPower

    If (SysPower.BatteryLifePercent <> iPowerLevel) Or ((SysPower.ACLineStatus <> 0) <> bAttachedToAC) Then

        bAttachedToAC = (SysPower.ACLineStatus <> 0)
        
        iPowerLevel = SysPower.BatteryLifePercent

        UpdateTrayToolTip

        If bShowDesktopIcon = True Then
            DesktopBattery.PaintImage
        End If

    End If

    If SysPower.ACLineStatus = 0 Then
        nID.hIcon = IL.ListImages(SysPower.BatteryLifePercent).Picture

     ElseIf (SysPower.BatteryLifePercent < 100) Then
        nID.hIcon = IL.ListImages(SysPower.BatteryLifePercent).Picture

        Select Case Animate

         Case 2
            X = Second(Now) Mod 8

            If X = 0 Then Y = 0

            If X < 5 And Y < 4 Then
                Y = Y + 1
                nID.hIcon = IL.ListImages("Ch" & Y).Picture
            End If

         Case 1
            X = Second(Now) Mod 5

            If X = 0 Then
                nID.hIcon = IL.ListImages("Bolt").Picture
            End If

        End Select

     Else
        nID.hIcon = IL.ListImages("AC").Picture

    End If

    If Tray Then
        Shell_NotifyIcon NIM_MODIFY, nID

     Else

        If Shell_NotifyIcon(NIM_ADD, nID) = 1 Then
            Tray = True

         Else
            MsgBox "Unable to load System Tray Icon, " & "Battery Monitor will close now!", vbExclamation, "Error..."
            Unload Me

        End If

    End If

End Sub

Private Sub UpdateInterval()

    UpdateKey HKEY_LOCAL_MACHINE, RegKey, "Interval", Timer.Interval
    UpdateKey HKEY_LOCAL_MACHINE, RegKey, "Animate", Str(Animate)

End Sub

Private Sub UpdateTrayToolTip()

    sTooltipInfo = " " & sAppName & " " & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine
    sTooltipInfo = sTooltipInfo & " Battery    : " & SysPower.BatteryLifePercent & "%" & vbNewLine
    sTooltipInfo = sTooltipInfo & " AC Status: " & IIf(SysPower.ACLineStatus = 1, "On AC", "Battery") & vbNullChar

    nID.szTip = sTooltipInfo

End Sub

Private Sub Validate(FileName As String)

    FileName = Replace(FileName, "\\", "\")

    If InStr(FileName, " ") > 0 Then
        FileName = Chr(34) & FileName & Chr(34)
    End If

End Sub

