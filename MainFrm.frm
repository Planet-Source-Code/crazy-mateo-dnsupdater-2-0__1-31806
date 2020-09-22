VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MainFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DNSUpdater"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame0 
      Height          =   3135
      Index           =   6
      Left            =   11040
      TabIndex        =   71
      Top             =   480
      Width           =   4935
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "#insults @ irc.atomicchat.net"
         Height          =   255
         Index           =   10
         Left            =   2160
         TabIndex        =   82
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "110684275"
         Height          =   255
         Index           =   8
         Left            =   2160
         TabIndex        =   81
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.crazymateo.com"
         Height          =   255
         Index           =   6
         Left            =   2160
         TabIndex        =   80
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Webmaster@crazymateo.com"
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   79
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "IRC:"
         Height          =   255
         Index           =   9
         Left            =   1200
         TabIndex        =   78
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "ICQ:"
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   77
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Homepage:"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   76
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   75
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image6 
         BorderStyle     =   1  'Fixed Single
         Height          =   2775
         Left            =   120
         Picture         =   "MainFrm.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crazy Mateo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   74
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "By"
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
         Index           =   1
         Left            =   1200
         TabIndex        =   73
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DNSUpdater 2.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   72
         Top             =   360
         Width           =   3615
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   1200
         X2              =   4800
         Y1              =   1545
         Y2              =   1545
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   1200
         X2              =   4800
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin VB.Frame Frame0 
      Height          =   3135
      Index           =   0
      Left            =   240
      TabIndex        =   66
      Top             =   480
      Width           =   4935
      Begin VB.CommandButton Command0 
         Caption         =   "Quit"
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   106
         Top             =   2760
         WhatsThisHelpID =   1104
         Width           =   855
      End
      Begin VB.Timer Timer0 
         Interval        =   1000
         Left            =   1080
         Top             =   840
      End
      Begin MSWinsockLib.Winsock Winsock0 
         Left            =   600
         Top             =   840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin InetCtlsObjects.Inet Inet0 
         Left            =   720
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.PictureBox Picture0 
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   1320
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   96
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComctlLib.ImageList ImageList0 
         Left            =   120
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MainFrm.frx":6D1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MainFrm.frx":6E74
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MainFrm.frx":7886
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MainFrm.frx":79E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MainFrm.frx":7B3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MainFrm.frx":7C94
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MainFrm.frx":7FAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MainFrm.frx":8548
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MainFrm.frx":8AE2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MainFrm.frx":8D3C
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MainFrm.frx":8F96
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MainFrm.frx":91F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MainFrm.frx":944A
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MainFrm.frx":96A4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog CommonDialog0 
         Left            =   120
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView ListView0 
         Height          =   2295
         Left            =   120
         TabIndex        =   105
         Top             =   220
         WhatsThisHelpID =   1000
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton Command0 
         Caption         =   "Sync"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   70
         Top             =   2760
         WhatsThisHelpID =   1101
         Width           =   855
      End
      Begin VB.CommandButton Command0 
         Caption         =   "Help"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   69
         Top             =   2760
         WhatsThisHelpID =   1102
         Width           =   855
      End
      Begin VB.CommandButton Command0 
         Caption         =   "Hide"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   68
         Top             =   2760
         WhatsThisHelpID =   1103
         Width           =   855
      End
      Begin VB.CommandButton Command0 
         Caption         =   "Update"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   67
         Top             =   2760
         WhatsThisHelpID =   1100
         Width           =   855
      End
      Begin VB.Line Line0 
         X1              =   120
         X2              =   4800
         Y1              =   2640
         Y2              =   2640
      End
   End
   Begin VB.Frame Frame0 
      Height          =   3135
      Index           =   4
      Left            =   5640
      TabIndex        =   57
      Top             =   3840
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   102
         Top             =   1320
         WhatsThisHelpID =   1341
         Width           =   855
      End
      Begin VB.ListBox List4 
         Height          =   840
         Left            =   840
         TabIndex        =   89
         Top             =   240
         WhatsThisHelpID =   1640
         Width           =   3975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Edit"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   88
         Top             =   840
         WhatsThisHelpID =   1142
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Delete"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   87
         Top             =   540
         WhatsThisHelpID =   1141
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Add"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   86
         Top             =   240
         WhatsThisHelpID =   1140
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   84
         Top             =   1320
         WhatsThisHelpID =   1340
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Add"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   60
         Top             =   1800
         WhatsThisHelpID =   1143
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Delete"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   59
         Top             =   1800
         WhatsThisHelpID =   1144
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Edit"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   58
         Top             =   1800
         WhatsThisHelpID =   1145
         Width           =   615
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   855
         Left            =   120
         TabIndex        =   61
         Top             =   2160
         WhatsThisHelpID =   1040
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   1508
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Keyword:"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   101
         Top             =   1320
         WhatsThisHelpID =   1341
         Width           =   735
      End
      Begin VB.Line Line4 
         Index           =   1
         X1              =   4800
         X2              =   120
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   83
         Top             =   1320
         WhatsThisHelpID =   1340
         Width           =   975
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   4800
         X2              =   120
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin VB.Frame Frame0 
      Height          =   3135
      Index           =   5
      Left            =   5640
      TabIndex        =   47
      Top             =   7200
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox Text5 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   960
         TabIndex        =   64
         Top             =   1080
         WhatsThisHelpID =   1351
         Width           =   3495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Edit"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   56
         Top             =   1560
         WhatsThisHelpID =   1152
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Delete"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   55
         Top             =   1560
         WhatsThisHelpID =   1151
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   54
         Top             =   1560
         WhatsThisHelpID =   1150
         Width           =   615
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   1095
         Left            =   120
         TabIndex        =   53
         Top             =   1920
         WhatsThisHelpID =   1050
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   1931
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Sounds"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   52
         Top             =   240
         WhatsThisHelpID =   1551
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Pictures"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Value           =   -1  'True
         WhatsThisHelpID =   1550
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   960
         TabIndex        =   50
         Top             =   600
         WhatsThisHelpID =   1350
         Width           =   3495
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "MainFrm.frx":98FE
         Left            =   2040
         List            =   "MainFrm.frx":990E
         TabIndex        =   48
         Text            =   "Default"
         Top             =   240
         WhatsThisHelpID =   1450
         Width           =   2775
      End
      Begin VB.Line Line5 
         Index           =   1
         X1              =   4800
         X2              =   120
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Image Image5 
         Height          =   240
         Index           =   1
         Left            =   4560
         Picture         =   "MainFrm.frx":9935
         Stretch         =   -1  'True
         Top             =   1080
         WhatsThisHelpID =   1751
         Width           =   240
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Log File:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   65
         Top             =   1080
         WhatsThisHelpID =   1351
         Width           =   735
      End
      Begin VB.Image Image5 
         Height          =   240
         Index           =   0
         Left            =   4560
         Picture         =   "MainFrm.frx":9D77
         Stretch         =   -1  'True
         Top             =   600
         WhatsThisHelpID =   1750
         Width           =   240
      End
      Begin VB.Line Line5 
         Index           =   0
         X1              =   4800
         X2              =   120
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Target:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   600
         WhatsThisHelpID =   1350
         Width           =   615
      End
   End
   Begin VB.Frame Frame0 
      Height          =   3135
      Index           =   3
      Left            =   5640
      TabIndex        =   42
      Top             =   480
      Visible         =   0   'False
      Width           =   4935
      Begin VB.ListBox List3 
         Height          =   840
         Left            =   840
         TabIndex        =   85
         Top             =   240
         WhatsThisHelpID =   1630
         Width           =   3975
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1695
         Left            =   120
         TabIndex        =   46
         Top             =   1320
         WhatsThisHelpID =   1030
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Edit"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   45
         Top             =   840
         WhatsThisHelpID =   1132
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   540
         WhatsThisHelpID =   1131
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   240
         WhatsThisHelpID =   1130
         Width           =   615
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   4800
         X2              =   120
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin VB.Frame Frame0 
      Height          =   3135
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   7200
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   1200
         TabIndex        =   103
         Top             =   2400
         WhatsThisHelpID =   1329
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   3480
         PasswordChar    =   "*"
         TabIndex        =   98
         Top             =   1920
         WhatsThisHelpID =   1328
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   3960
         TabIndex        =   40
         Top             =   1080
         WhatsThisHelpID =   1324
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   37
         Top             =   1920
         WhatsThisHelpID =   1327
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1200
         TabIndex        =   33
         Top             =   600
         WhatsThisHelpID =   1321
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   4080
         TabIndex        =   32
         Text            =   "80"
         Top             =   600
         WhatsThisHelpID =   1322
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   3840
         PasswordChar    =   "*"
         TabIndex        =   31
         Top             =   240
         WhatsThisHelpID =   1320
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "MainFrm.frx":A1B9
         Left            =   1200
         List            =   "MainFrm.frx":A1BB
         TabIndex        =   26
         Top             =   240
         WhatsThisHelpID =   1420
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1200
         TabIndex        =   25
         Text            =   "http://"
         Top             =   1080
         WhatsThisHelpID =   1323
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Router"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Value           =   -1  'True
         WhatsThisHelpID =   1520
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "HTTP Page"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   23
         Top             =   2760
         WhatsThisHelpID =   1522
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   1200
         TabIndex        =   22
         Top             =   1560
         WhatsThisHelpID =   1325
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   4080
         TabIndex        =   21
         Text            =   "6667"
         Top             =   1560
         WhatsThisHelpID =   1326
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "IRC Server"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   20
         Top             =   2760
         WhatsThisHelpID =   1521
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Proxy:"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   104
         Top             =   2400
         WhatsThisHelpID =   1329
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name:"
         Height          =   255
         Index           =   9
         Left            =   2640
         TabIndex        =   97
         Top             =   1920
         WhatsThisHelpID =   1328
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Keyword:"
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   39
         Top             =   1080
         WhatsThisHelpID =   1324
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nick Name:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   38
         Top             =   1920
         WhatsThisHelpID =   1327
         Width           =   1095
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   4800
         X2              =   120
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   4800
         X2              =   120
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   600
         WhatsThisHelpID =   1321
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   35
         Top             =   600
         WhatsThisHelpID =   1322
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   34
         Top             =   240
         WhatsThisHelpID =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Router:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   240
         WhatsThisHelpID =   1420
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "HTTP Page:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         WhatsThisHelpID =   1323
         Width           =   1095
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   120
         X2              =   4800
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "IRC Server:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         WhatsThisHelpID =   1325
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   255
         Index           =   7
         Left            =   3600
         TabIndex        =   27
         Top             =   1560
         WhatsThisHelpID =   1326
         Width           =   495
      End
   End
   Begin VB.Frame Frame0 
      Height          =   3135
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   4935
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   3480
         TabIndex        =   99
         Top             =   2280
         WhatsThisHelpID =   1411
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   3480
         PasswordChar    =   "*"
         TabIndex        =   95
         Top             =   2680
         WhatsThisHelpID =   1315
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   1080
         TabIndex        =   93
         Top             =   2680
         WhatsThisHelpID =   1314
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   91
         Top             =   2280
         WhatsThisHelpID =   1410
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Auto Sync"
         Height          =   255
         Index           =   11
         Left            =   3600
         TabIndex        =   63
         Top             =   960
         WhatsThisHelpID =   12111
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Logging"
         Height          =   255
         Index           =   7
         Left            =   1920
         TabIndex        =   62
         Top             =   960
         WhatsThisHelpID =   1217
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Close Programs"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   41
         Top             =   960
         WhatsThisHelpID =   1213
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Play Sounds"
         Height          =   255
         Index           =   6
         Left            =   1920
         TabIndex        =   18
         Top             =   720
         WhatsThisHelpID =   1216
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Animate Tray"
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   17
         Top             =   240
         WhatsThisHelpID =   1214
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Auto Quit"
         Height          =   255
         Index           =   8
         Left            =   3600
         TabIndex        =   16
         Top             =   240
         WhatsThisHelpID =   1218
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Programs"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   720
         WhatsThisHelpID =   1212
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Auto Retry"
         Height          =   255
         Index           =   9
         Left            =   3600
         TabIndex        =   14
         Top             =   480
         WhatsThisHelpID =   1219
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4440
         MaxLength       =   2
         TabIndex        =   12
         Text            =   "1"
         Top             =   1800
         WhatsThisHelpID =   1313
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "30"
         Top             =   1800
         WhatsThisHelpID =   1311
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   4440
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "1"
         Top             =   1440
         WhatsThisHelpID =   1312
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "LAN Connect"
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   6
         Top             =   480
         WhatsThisHelpID =   1215
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "1"
         Top             =   1440
         WhatsThisHelpID =   1310
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Auto Update"
         Height          =   255
         Index           =   10
         Left            =   3600
         TabIndex        =   4
         Top             =   720
         WhatsThisHelpID =   12110
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Start Hidden"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   480
         WhatsThisHelpID =   1211
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Run On Startup"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         WhatsThisHelpID =   1210
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   255
         Index           =   7
         Left            =   2520
         TabIndex        =   100
         Top             =   2685
         WhatsThisHelpID =   1315
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   94
         Top             =   2685
         WhatsThisHelpID =   1314
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Default IP:"
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   92
         Top             =   2280
         WhatsThisHelpID =   1411
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Service:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   90
         Top             =   2280
         WhatsThisHelpID =   1410
         Width           =   855
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   4800
         X2              =   120
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   120
         X2              =   4800
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2400
         X2              =   2400
         Y1              =   1440
         Y2              =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Number Of Retries:"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   13
         Top             =   1800
         WhatsThisHelpID =   1313
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Retry Delay"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         WhatsThisHelpID =   1311
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Timer Interval:"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   9
         Top             =   1440
         WhatsThisHelpID =   1312
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Update Delay:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         WhatsThisHelpID =   1310
         Width           =   1575
      End
   End
   Begin MSComctlLib.TabStrip TabStrip0 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   1800
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6376
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Status"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Setup"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "LAN Connect"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Routers"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Services"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Misc"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private UpdateInfo(4) As Integer
Private CountInfo(4) As Integer

Private Sub Check1_Click(Index As Integer)
On Error GoTo err_handler
Dim x As Integer
WriteToDatabase DataPath, "Setup", "Check" & Index, "Value", Check1(Index).value, ";"
Select Case Index
    Case 0
        If Check1(0).value = 1 Then
            If InRun("DNSUpdater") = False Then AddToRun "DNSUpdater", App.path & "\DNSUpdater.exe"
        Else
            If InRun("DNSUpdater") = True Then DeleteFromRun "DNSUpdater"
        End If
    Case 5
        If TabStrip0.Tabs(2).Tag = True Then
            If Check1(5).value = 1 Then
                UpdateInfo(2) = UpdateInfo(2) * 60
            Else
                UpdateInfo(2) = UpdateInfo(2) / 60
            End If
        End If
    Case 10
        UpdateInfo(4) = IIf(Check1(Index).value = 0, 0, 1)
End Select
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Save Setting."
End Sub

Private Sub Combo1_Change(Index As Integer)
On Error Resume Next: Combo1_Click (Index)
End Sub

Private Sub Combo1_Click(Index As Integer)
On Error GoTo err_handler
If Index = 0 Then
    If IsName(Combo1(0), service) Then WriteToDatabase DataPath, "Setup", "Combo" & Index, "Value", Combo1(0), ";"
ElseIf Index = 1 Then
    If ValidIP(GetFullIP(Combo1(1))) Then WriteToDatabase DataPath, "Setup", "Combo" & Index, "Value", Combo1(1), ";"
End If
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Save Default IP."
End Sub

Private Sub Combo1_DropDown(Index As Integer)
On Error GoTo err_handler
Dim temp As String
temp = Combo1(1)
If Index = 1 Then ListAddString Combo1(1), GetLocalIPs, ";"
Combo1(1) = temp
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Refresh IP list."
End Sub

Private Sub Combo2_Click()
On Error Resume Next: If IsName(Combo2, Router) Then WriteToDatabase DataPath, "LANConnect", "Combo", "Value", Combo2, ";"
End Sub

Private Sub Combo5_Click()
On Error GoTo err_handler
Dim temp As String
temp = ReadFromDatabase(DataPath, "Misc", IIf(Option5(0).value = True, "Picture", "Sound") & Combo5, "Value", ";", "norecord", "null", True)
If temp = "null" Or temp = "norecord" Or PathFileExists(ApplyConstants(temp)) = 0 Then Text5(0) = vbNullString Else: Text5(0) = temp
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Load " & IIf(Option5(0).value = True, "Picture", "Sound") & "."
End Sub

Public Sub Command0_Click(Index As Integer)
On Error GoTo err_handler
Dim temp As String
Dim x As Integer
Select Case Index
    Case 0
        AddMessage Status, "Manual Update"
        CountInfo(0) = UpdateInfo(0)
    Case 1
        temp = ReadFromDatabase(DataPath, "Setup", "Check4", "Value", ";", "norecord", "null", True, False)
        RunSynchronize Me, DataPath, IIf(temp = "1", True, False), playsounds
    Case 2
        OpenHelp Me, App.path & "\help.chm"
    Case 3
        Visible = IIf(Visible = True, False, True)
        If Visible = False Then UnloadVisibleForms
        If Menu <> 0 Then
            ModifyMenu Menu, 0, MF_STRING, 0, CStr(IIf(Visible = True, "Hide", "Show"))
            SetMenuItemBitmaps Menu, 0, 0, ImageList0.ListImages(IIf(Visible = True, 10, 9)).Picture, ImageList0.ListImages(IIf(Visible = True, 10, 9)).Picture
        End If
    Case 4
        Unload Me
End Select
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Execute Command."
End Sub

Private Sub Command3_Click(Index As Integer)
On Error GoTo err_handler
Dim temp As String
If List3.ListIndex = -1 And Index <> 0 Then MsgBox "You Must Select A Router.": Exit Sub
If Index = 1 Then
    DeleteFromDatabase DataPath, "Routers", List3
    List3.RemoveItem List3.ListIndex
    If List3.ListCount <> 0 Then List3.Selected(0) = True: LoadRouterOptions Else: ListView3.ListItems.Clear
    temp = GetTableNames(DataPath, "Routers", ";")
    ListAddString Combo2, IIf(temp = "norecords", vbNullString, temp), ";"
    LoadRouter
Else
    Modification = name & ";List3;Routers;" & IIf(Index = 0, "0;Add Router", "1;Edit Router")
    If IsFormLoaded("ModificationFrm") = True Then Unload ModificationFrm
    ModificationFrm.Show vbModal
End If
FixHeaders ListView4, 6
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Modify Router."
End Sub

Private Sub Command4_Click(Index As Integer)
On Error GoTo err_handler
Dim temp As String
Dim replacestr As String
Select Case Index
    Case 0 To 2
        If List4.ListIndex = -1 And Index <> 0 Then MsgBox "You Must Select A Service.": Exit Sub
        If Index = 1 Then
            DeleteFromDatabase DataPath, "Services", List4
            List4.RemoveItem List4.ListIndex
            If List4.ListCount <> 0 Then List4.Selected(0) = True: LoadServiceOptions Else: ListView4.ListItems.Clear: Text4(0) = vbNullString: Text4(1) = vbNullString
            temp = GetTableNames(DataPath, "Services", ";")
            ListAddString Combo1(0), IIf(temp = "norecords", vbNullString, temp), ";"
            LoadService
        Else
            Modification = name & ";List4;Services;" & IIf(Index = 0, "3;Add Service", "4;Edit Service")
        End If
    Case 3 To 5
        If ListView4.ListItems.Count = 0 And Index <> 3 Or List4.ListIndex = -1 Then MsgBox "You Must Select A Field.": Exit Sub
        If Index = 4 Then
            temp = ReadFromDatabase(DataPath, "Services", List4, "Fields", ";", "norecord", "null", False)
            If temp = "norecord" Then AddError MISSING_DATA, "Unable To Remove Service Field.": Exit Sub
            If temp = "null" Then AddError MISSING_DATA, "Unable Remove Service Field.": Exit Sub
            If ListView4.SelectedItem.Index <> 1 Then replacestr = "&"
            replacestr = replacestr & ListView4.ListItems(ListView4.SelectedItem.Index) & "=" & ListView4.ListItems(ListView4.SelectedItem.Index).ListSubItems(1)
            If ListView4.SelectedItem.Index = 1 And ListView4.ListItems.Count <> 1 Then replacestr = replacestr & "&"
            temp = Replace(temp, replacestr, "")
            WriteToDatabase DataPath, "Services", List4, "Fields", temp, ";"
            ListView4.ListItems.Remove ListView4.SelectedItem.Index
        Else
            Modification = name & ";ListView4;Services;" & IIf(Index = 3, "5;Add Field", "6;Edit Field") & ";List4"
        End If
End Select
If Index <> 1 And Index <> 4 Then
    If IsFormLoaded("ModificationFrm") = True Then Unload ModificationFrm
    ModificationFrm.Show vbModal
End If
FixHeaders ListView4, 2
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Modify Service."
End Sub

Private Sub Command5_Click(Index As Integer)
On Error GoTo err_handler
If ListView5.ListItems.Count = 0 And Index <> 0 Then MsgBox "You Must Select A Program.": Exit Sub
If Index = 1 Then
    DeleteFromDatabase DataPath, "Misc", "Program" & ListView5.ListItems(ListView5.SelectedItem.Index).Text
    ListView5.ListItems.Remove ListView5.SelectedItem.Index
Else
    Modification = name & ";ListView5;Misc;" & IIf(Index = 0, "7;Add Program", "8;Edit Program")
    If IsFormLoaded("ModificationFrm") = True Then Unload ModificationFrm
    ModificationFrm.Show vbModal
End If
FixHeaders ListView5, 3
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Modify Program."
End Sub

Private Sub Form_Load()
On Error GoTo err_handler
Dim temp As String
Width = 5490: Height = 4320
ListView0.ColumnHeaders(1).Width = ListView0.Width - 60
temp = ReadFromDatabase(DataPath, "Setup", "Check10", "Value", ";", "norecord", "null", True, False)
UpdateInfo(4) = IIf(temp = "1", 1, 0)
LoadUpdateInfo
temp = ReadFromDatabase(DataPath, "Setup", "Check5", "Value", ";", "norecord", "null", True, False)
If temp = "1" Then UpdateInfo(2) = UpdateInfo(2) * 60
If UpdateInfo(4) = 1 Then CountInfo(2) = UpdateInfo(2)
AddHeaders ListView3, "Property;Value", ";"
AddHeaders ListView4, "Name;Value", ";"
AddHeaders ListView5, "Name;Value", ";"
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Complete Main Startup."
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next: MainQuit
End Sub

Private Sub Image5_Click(Index As Integer)
On Error GoTo err_handler
CommonDialog0.InitDir = "C:\"
Select Case Index
    Case 0
        If Option5(0).value = True Then
            CommonDialog0.DialogTitle = "Browse Image Target"
            CommonDialog0.Filter = "Icon (*.ico)|*.ico|Bitmap (*.bmp)|*.bmp"
        Else
            CommonDialog0.DialogTitle = "Browse Sound Target"
            CommonDialog0.Filter = "WAV File (*.wav)|*.wav"
        End If
        CommonDialog0.ShowOpen
    Case 1
        CommonDialog0.DialogTitle = "Save Log File"
        CommonDialog0.Filter = "Log File (*.log)|*.log"
        CommonDialog0.ShowSave
End Select
If CommonDialog0.filename = vbNullString Then Exit Sub
Text5(Index) = CommonDialog0.filename: CommonDialog0.filename = vbNullString
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Browse."
End Sub

Private Sub Inet0_StateChanged(ByVal State As Integer)
On Error Resume Next
Dim ReturnString As String
If State = icError Then Inet0.Tag = "error": Exit Sub
ReturnString = Inet0.GetChunk(2048, icString)
Do Until Len(ReturnString) = 0
    Inet0.Tag = Inet0.Tag & ReturnString
    ReturnString = Inet0.GetChunk(2048, icString)
    DoEvents
Loop
End Sub

Private Sub List3_Click()
On Error Resume Next: LoadRouterOptions
End Sub

Private Sub List3_DblClick()
On Error Resume Next: Command3_Click 2
End Sub

Private Sub List4_Click()
On Error Resume Next: LoadServiceOptions
End Sub

Private Sub List4_DblClick()
On Error Resume Next: Command4_Click 2
End Sub

Private Sub ListView0_Click()
On Error Resume Next: ListView0.ListItems(1).Selected = True: ListView0.ListItems(ListView0.SelectedItem.Index).Selected = False
End Sub

Private Sub ListView0_GotFocus()
On Error Resume Next: ListView0_Click
End Sub

Private Sub ListView0_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next: ListView0_Click
End Sub

Private Sub ListView0_LostFocus()
On Error Resume Next: ListView0_Click
End Sub

Private Sub ListView3_DblClick()
On Error GoTo err_handler
If ListView3.ListItems.Count = 0 Then MsgBox "Nothing To Edit.": Exit Sub
Modification = name & ";ListView3;Routers;2;Edit Property;List3"
If IsFormLoaded("ModificationFrm") = True Then Unload ModificationFrm
ModificationFrm.Show vbModal
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Edit Router Property."
End Sub

Private Sub ListView4_DblClick()
On Error Resume Next: Command4_Click 5
End Sub

Private Sub ListView5_DblClick()
On Error Resume Next: Command5_Click 2
End Sub

Private Sub Option2_Click(Index As Integer)
On Error Resume Next: WriteToDatabase DataPath, "LANConnect", "Option", "Value", CStr(Index), ";"
End Sub

Private Sub Option5_Click(Index As Integer)
On Error GoTo err_handler
ListAddString Combo5, IIf(Index = 0, "Default;Updating;Synchronizing;Retrying;Success;Failure", "Start;Quit;Updating;Synchronizing;Retrying;Success;Failure")
If Combo5.ListCount <> 0 Then Combo5.ListIndex = 0: Combo5_Click
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Load List."
End Sub

Private Sub Picture0_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo err_handler
Dim cursorpos As POINTAPI
Msg = x / Screen.TwipsPerPixelX
If Msg = WM_LBUTTONUP Then
    Command0_Click 3
ElseIf Msg = WM_RBUTTONUP Then
    GetCursorPos cursorpos
    TrackPopupMenu Menu, TPM_BOTTOMALIGN, cursorpos.x, cursorpos.y, 0, hwnd, ByVal 0&
End If
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Execute Tray Command."
End Sub

Private Sub DisplayFrame(Index As Integer)
On Error GoTo err_handler
Dim x As Integer
If Index < 0 Or Index > TabStrip0.Tabs.Count - 1 Then AddError MISSING_DATA, "Failed To Display Frame.": Exit Sub
For x = 0 To Frame0.Count - 1
    If x = Index Then
        Frame0(x).Visible = True
        Frame0(x).Left = 240
        Frame0(x).Top = 480
    Else
        Frame0(x).Visible = False
    End If
Next x
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Display Frame."
End Sub

Private Sub TabStrip0_Mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo err_handler
Dim temp As String
If TabStrip0.Tabs(TabStrip0.SelectedItem.Index).Tag <> True Then
    Select Case TabStrip0.SelectedItem.Index
        Case 2
            temp = GetTableNames(DataPath, "Services", ";")
            ListAddString Combo1(0), IIf(temp = "norecords", vbNullString, temp), ";"
            ListAddString Combo1(1), GetLocalIPs, ";"
            LoadSettings Me, DataPath, Setup, "Setup"
        Case 3
            temp = GetTableNames(DataPath, "Routers", ";")
            ListAddString Combo2, IIf(temp = "norecords", vbNullString, temp), ";"
            LoadSettings Me, DataPath, LANConnect, "LANConnect"
        Case 4
            temp = GetTableNames(DataPath, "Routers", ";")
            ListAddString List3, IIf(temp = "norecords", vbNullString, temp), ";"
            If List3.ListCount <> 0 Then List3.Selected(0) = True: LoadRouterOptions
        Case 5
            temp = GetTableNames(DataPath, "Services", ";")
            ListAddString List4, IIf(temp = "norecords", vbNullString, temp), ";"
            If List4.ListCount <> 0 Then List4.Selected(0) = True: LoadServiceOptions
        Case 6
            If Combo5.ListCount <> 0 Then Combo5.ListIndex = 0
            temp = ReadFromDatabase(DataPath, "Misc", "LogFile", "Value", ";", "norecord", "null", True)
            If temp = "norecord" Or temp = "null" Then Text5(1) = vbNullString Else: Text5(1) = temp
            LoadPrograms
    End Select
    TabStrip0.Tabs(TabStrip0.SelectedItem.Index).Tag = True
End If
DisplayFrame TabStrip0.SelectedItem.Index - 1
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failure Loading Tab Settings."
End Sub

Private Sub Text1_Change(Index As Integer)
On Error GoTo err_handler
Dim value As String
Select Case Index
    Case 0
        value = IIf(Not IsNumeric(Text1(Index)) Or Text1(Index) < 1, 1, Text1(Index))
    Case 1
        value = IIf(Not IsNumeric(Text1(Index)) Or Text1(Index) < 30, 30, Text1(Index))
    Case 2 To 3
        value = IIf(Not IsNumeric(Text1(Index)) Or Text1(Index) < 1, 1, Text1(Index))
    Case Else
        value = Text1(Index)
End Select
If Index >= 0 And Index <= 3 Then UpdateInfo(Index) = value
WriteToDatabase DataPath, "Setup", "Text" & Index, "Value", value, ";"
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Change Value."
End Sub

Private Sub Text2_Change(Index As Integer)
On Error GoTo err_handler
Dim value As String
Select Case Index
    Case 1
        value = IIf(ValidIP(Text2(Index)) = False, "0.0.0.0", Text2(Index))
    Case 2
        value = IIf(Not IsNumeric(Text2(Index)) Or Text2(Index) <= 0, 80, Text2(Index))
    Case 3
        value = IIf(Not Text2(Index) Like "*://*", "http://", Text2(Index))
    Case 5, 9
        value = IIf(ValidIP(Text2(Index)) = False And Not Text2(Index) Like "*.*", vbNullString, Text2(Index))
    Case 6
        value = IIf(Not IsNumeric(Text2(Index)) Or Text2(Index) < 1, 6667, Text2(Index))
    Case Else
        value = Text2(Index)
End Select
WriteToDatabase DataPath, "LANConnect", "Text" & Index, "Value", value, ";"
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Change Value."
End Sub

Public Sub LoadRouterOptions()
On Error GoTo err_handler
Dim temp() As String
Dim x As Integer
If List3.ListIndex = -1 Then MsgBox "You Must Select A Router.": Exit Sub
temp = Split(ReadFromDatabase(DataPath, "Routers", List3, "Name;LogIn;LogOut;Status;Keyword", ";", "norecord", "null", False), ";")
If temp(0) = "norecord" Then AddError MISSING_DATA, "Unable To Load Router.": Exit Sub
If UBound(temp) <> 4 Then GoTo err_handler
For x = 0 To UBound(temp)
    If LCase(temp(x)) = "null" Then temp(x) = vbNullString
    If x >= 1 And x <= 3 And Not temp(x) Like "*://*" Then temp(x) = "http://"
Next x
ListViewAddString ListView3, "Name^" & temp(0) & ";Log In^" & temp(1) & ";Log Out^" & temp(2) & ";Status^" & temp(3) & ";Keyword^" & temp(4), ";", "^"
FixHeaders ListView4, 6
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Load Router."
End Sub

Public Sub LoadRouter()
On Error GoTo err_handler
Dim temp As String
temp = GetTableNames(DataPath, "Routers", ";")
ListAddString Combo2, IIf(temp = "norecords", vbNullString, temp), ";"
temp = ReadFromDatabase(DataPath, "LANConnect", "Combo", "Value", ";", "norecord", "null", True)
Combo2 = IIf(Not IsName(temp, Router) Or temp = "norecord" Or temp = "null", vbNullString, temp)
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Load Router Option."
End Sub

Public Sub RouterNameEdit(oldv As String, newv As String)
On Error Resume Next: If LCase(oldv) = LCase(Combo2) Then Combo2 = newv: Combo2_Click
End Sub

Public Sub LoadServiceOptions()
On Error GoTo err_handler
Dim temp() As String
If List4.ListIndex = -1 Then MsgBox "You Must Select A Service.": Exit Sub
temp = Split(ReadFromDatabase(DataPath, "Services", List4, "Address;Fields;Keyword", ";", "norecord", "null", False), ";")
If UBound(temp) <> 2 Then AddError MISSING_DATA, "Failed To Load Service.": Exit Sub
Text4(0) = IIf(Not temp(0) Like "*://*", "http://", temp(0)): ListViewAddString ListView4, temp(1), "&", "="
Text4(1) = IIf(temp(2) = "null" Or temp(2) = vbNullString, vbNullString, temp(2))
FixHeaders ListView4, 2
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Load Service."
End Sub

Public Sub LoadService()
On Error GoTo err_handler
Dim temp As String
temp = GetTableNames(DataPath, "Services", ";")
ListAddString Combo1(0), IIf(temp = "norecords", vbNullString, temp), ";"
temp = ReadFromDatabase(DataPath, "Setup", "Combo", "Value", ";", "norecord", "null", True)
Combo1(0) = IIf(Not IsName(temp, service) Or temp = "norecord" Or temp = "null", vbNullString, temp)
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Load Service Option."
End Sub

Public Sub ServiceNameEdit(oldv As String, newv As String)
On Error Resume Next: If LCase(oldv) = LCase(Combo1(0)) Then Combo1(0) = newv: Combo1_Click 0
End Sub

Public Sub LoadPrograms()
On Error GoTo err_handler
Dim temp1() As String
Dim temp2, temp3 As String
Dim x As Integer
temp1 = Split(GetTableNames(DataPath, "Misc", ";"), ";")
For x = 0 To UBound(temp1)
    If temp1(x) Like "Program*" Then
        temp2 = ReadFromDatabase(DataPath, "Misc", temp1(x), "Value", ";", "norecord", "null", False)
        If temp2 = "norecord" Then GoTo err_handler
        If temp2 <> "null" And LCase(Right(temp2, 4)) = ".exe" Then temp3 = IIf(temp3 = vbNullString, vbNullString, temp3 & ";") & Mid(temp1(x), 8, Len(temp1(x)) - 7) & "," & temp2
    End If
Next x
ListViewAddString ListView5, temp3, ";", ","
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Load Programs."
End Sub

Public Sub LoadUpdateInfo()
On Error GoTo err_handler
Dim temp As String
Dim x As Integer
For x = 0 To 3
    temp = ReadFromDatabase(DataPath, "Setup", "Text" & x, "Value", ";", "norecord", "null", True, False)
    Select Case x
        Case 0
            If IsNumeric(temp) Then UpdateInfo(x) = CInt(IIf(CInt(temp) < 1, 1, temp)) Else UpdateInfo(x) = 1
        Case 1
            If IsNumeric(temp) Then UpdateInfo(x) = CInt(IIf(CInt(temp) < 5, 5, temp)) Else UpdateInfo(x) = 5
        Case 2 To 3
            If IsNumeric(temp) Then UpdateInfo(x) = CInt(IIf(CInt(temp) < 1, 1, temp)) Else UpdateInfo(x) = 1
        Case Else
            If IsNumeric(temp) Then UpdateInfo(x) = CInt(temp) Else UpdateInfo(x) = 1
    End Select
Next x
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Load Update Info."
End Sub

Private Sub Text4_Change(Index As Integer)
On Error GoTo err_handler
If List4.ListIndex = -1 Then Exit Sub
If InString(Text4(Index), ";") Then MsgBox "Invalid Character Found.": Exit Sub
Select Case Index
    Case 0
        WriteToDatabase DataPath, "Services", List4, "Address", IIf(Not Text4(Index) Like "*://*", "http://", Text4(Index)), ";"
    Case 1
        WriteToDatabase DataPath, "Services", List4, "Keyword", Text4(Index), ";"
End Select
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Change Service Option."
End Sub

Private Sub Text5_Change(Index As Integer)
On Error GoTo err_handler
Select Case Index
    Case 0
        If PathFileExists(ApplyConstants(Text5(0))) = 0 And Trim(Text5(0)) <> vbNullString Or Not Right(Text5(0), 4) Like ".???" And Trim(Text5(0)) <> vbNullString Then AddError INVALID_TYPE, "Failed To Change Path.": Exit Sub
        If Combo5.ListIndex <> -1 Then
            WriteToDatabase DataPath, "Misc", IIf(Option5(0).value = True, "Picture", "Sound") & Combo5, "Value", Text5(Index), ";"
            If Option5(0).value = True And LCase(Combo5) = "default" Then ChangeTrayIcon pic, "Default", Picture0, ImageList0, Me
        End If
    Case 1
        If Not ApplyConstants(Text5(1)) Like "?:\*.log" Then Exit Sub
        WriteToDatabase DataPath, "Misc", "LogFile", "Value", Text5(1), ";"
End Select
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Change Path."
End Sub

Private Sub Timer0_Timer()
On Error GoTo err_handler
Dim temp(1) As String
If CountInfo(1) <> 0 Then
    If CountInfo(1) = UpdateInfo(1) Then
        If CountInfo(3) = UpdateInfo(3) Then
            CountInfo(1) = 0: CountInfo(3) = 0
        Else
            temp(0) = ReadFromDatabase(DataPath, "Setup", "Check4", "Value", ";", "norecord", "null", True, False)
            If RunUpdate(Me, IIf(temp(0) = "1", True, False), CountInfo(3) + 1, playsounds) = False Then
                CountInfo(1) = 1: CountInfo(3) = CountInfo(3) + 1
            Else
                CountInfo(1) = 0: CountInfo(3) = 0
                StartPrograms DataPath, "Misc"
            End If
        End If
    Else
        CountInfo(1) = CountInfo(1) + 1
    End If
ElseIf CountInfo(0) <> 0 Then
    If CountInfo(0) = UpdateInfo(0) Then
        temp(0) = ReadFromDatabase(DataPath, "Setup", "Check11", "Value", ";", "norecord", "null", True, False)
        temp(1) = ReadFromDatabase(DataPath, "Setup", "Check4", "Value", ";", "norecord", "null", True, False)
        If temp(0) = "1" Then RunSynchronize Me, DataPath, IIf(temp(1) = "1", True, False), playsounds
        If RunUpdate(Me, IIf(temp(1) = "1", True, False), , playsounds) = False Then
            temp(0) = ReadFromDatabase(DataPath, "Setup", "Check9", "Value", ";", "norecord", "null", True, False)
            If temp(0) = "1" Then CountInfo(1) = 1
        Else
            StartPrograms DataPath, "Misc"
        End If
        CountInfo(0) = 0
    Else
        CountInfo(0) = CountInfo(0) + 1
    End If
Else
    If UpdateInfo(4) = 1 Then
        If UpdateInfo(2) <= CountInfo(2) Then
            If NewIP(Me) = True Then
                If ExternalIP = "0.0.0.0" Then
                    ClosePrograms
                    CountInfo(2) = 0: Exit Sub
                End If
                AddMessage Status, "Start Delay"
                CountInfo(0) = 1
            End If
            CountInfo(2) = 0
        Else
            CountInfo(2) = CountInfo(2) + 1
        End If
    End If
End If
If MainMod.TrayIcon.szTip Like "DNSUpdater: Failure*" Or MainMod.TrayIcon.szTip Like "DNSUpdater: Success*" Or MainMod.TrayIcon.szTip Like "DNSUpdater: Synchronized*" Then
    If CountInfo(4) = 5 Then
        ChangeTrayIcon tip, "DNSUpdater"
        ChangeTrayIcon pic, "Default", Picture0, ImageList0, Me
        CountInfo(4) = 0
    Else
        CountInfo(4) = CountInfo(4) + 1
    End If
End If
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Unknown Error Occurred Running Timer."
End Sub

Private Sub Winsock0_Connect()
On Error GoTo err_handler
Dim temp(1) As String
Dim x As Integer
temp(0) = ReadFromDatabase(DataPath, "LANConnect", "Text7", "Value", ";", "norecord", "null", True, False)
temp(1) = ReadFromDatabase(DataPath, "LANConnect", "Text8", "Value", ";", "norecord", "null", True, False)
For x = 0 To 1
    If temp(x) = "norecord" Or temp(x) = "null" Then
        If x = 0 Then temp(x) = "DNSUpdater" & Int(Rnd * 10000)
        ElseIf x = 1 Then temp(x) = "DNSUpdater"
    End If
Next x
Winsock0.SendData "NICK " & temp(0) & vbCrLf
Winsock0.SendData "USER " & temp(1) & " " & Winsock0.LocalIP & " " & Winsock0.RemoteHostIP & " :" & temp(1) & vbCrLf
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Connect."
    Winsock0.Tag = "error1: Winsock0.Close"
End Sub

Private Sub Winsock0_DataArrival(ByVal bytesTotal As Long)
On Error GoTo err_handler
Dim temp As String
Winsock0.GetData temp, vbString
If InStr(temp, "001") <> 0 And temp Like "*@*" Then
    temp = Mid(temp, InStr(temp, "@") + 1)
    Winsock0.Tag = Mid(temp, 1, InStr(temp, vbNewLine) - 1)
    Winsock0.Close
ElseIf InStr(temp, "PING :") <> 0 Then
    Winsock0.SendData "pong " & Mid(temp, InStr(temp, "PING :") + 6)
End If
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Parse Arrival Data."
End Sub

Private Sub Winsock0_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next: Winsock0.Tag = "error2": Winsock0.Close
End Sub
