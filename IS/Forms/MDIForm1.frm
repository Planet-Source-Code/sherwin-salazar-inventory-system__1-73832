VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00F5F5F5&
   Caption         =   "InventorySystem"
   ClientHeight    =   10710
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14685
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer timerMonChild 
      Interval        =   1
      Left            =   5280
      Top             =   3600
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   4800
      Top             =   5760
   End
   Begin VB.PictureBox Picture4 
      Align           =   1  'Align Top
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      Height          =   60
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   14685
      TabIndex        =   11
      Top             =   990
      Width           =   14685
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   14685
      TabIndex        =   10
      Top             =   10215
      Width           =   14685
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2010-2011 by Salazar Software. All Right Reserved. E-mail: charminglang@yahoo.com : (+63) 09295218029"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   360
         TabIndex        =   18
         Top             =   195
         Width           =   10455
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   0
         Picture         =   "MDIForm1.frx":57E2
         Top             =   200
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   615
         Left            =   -240
         Picture         =   "MDIForm1.frx":5A26
         Stretch         =   -1  'True
         Top             =   -33
         Width           =   38175
      End
   End
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   5280
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   9165
      Left            =   0
      ScaleHeight     =   9165
      ScaleWidth      =   3045
      TabIndex        =   1
      Top             =   1050
      Width           =   3045
      Begin VB.PictureBox picMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   34
         ScaleHeight     =   4575
         ScaleWidth      =   2940
         TabIndex        =   2
         Top             =   1080
         Width           =   2940
         Begin MSComctlLib.ListView Listview1 
            Height          =   4260
            Left            =   30
            TabIndex        =   3
            Top             =   390
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   7514
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            OLEDragMode     =   1
            _Version        =   393217
            Icons           =   "i32x32"
            SmallIcons      =   "i32x32"
            ForeColor       =   -2147483641
            BackColor       =   -2147483634
            Appearance      =   0
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "MDIForm1.frx":5B47
            OLEDragMode     =   1
            NumItems        =   0
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   255
            Left            =   33
            TabIndex        =   12
            Top             =   180
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Select a Task"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   33
            TabIndex        =   4
            Top             =   -30
            Width           =   2055
         End
      End
      Begin lvButton.lvButtons_H cmdSet 
         Height          =   375
         Left            =   30
         TabIndex        =   8
         Top             =   5655
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   661
         Caption         =   "System Setings"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFHover         =   9069372
         cBhover         =   13016952
         LockHover       =   3
         cGradient       =   -2147483628
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483629
      End
      Begin lvButton.lvButtons_H cmdFile 
         Height          =   375
         Left            =   30
         TabIndex        =   9
         Top             =   675
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   661
         Caption         =   "Quick Launch"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFHover         =   9069372
         cBhover         =   13016952
         LockHover       =   3
         cGradient       =   -2147483628
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483629
      End
      Begin VB.PictureBox picSet 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   34
         ScaleHeight     =   3495
         ScaleWidth      =   2940
         TabIndex        =   5
         Top             =   6015
         Visible         =   0   'False
         Width           =   2940
         Begin MSComctlLib.ListView Listview2 
            Height          =   2745
            Left            =   30
            TabIndex        =   6
            Top             =   480
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   4842
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            OLEDragMode     =   1
            _Version        =   393217
            Icons           =   "ImageList1"
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483641
            BackColor       =   -2147483634
            Appearance      =   0
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "MDIForm1.frx":6421
            OLEDragMode     =   1
            NumItems        =   0
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   255
            Left            =   33
            TabIndex        =   17
            Top             =   180
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Select a Task"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   60
            TabIndex        =   7
            Top             =   -30
            Width           =   2055
         End
      End
      Begin VB.Label lblCurrentUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   750
         TabIndex        =   16
         Top             =   240
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome ,"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   165
         Left            =   45
         TabIndex        =   15
         Top             =   255
         Width           =   645
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   750
         TabIndex        =   14
         Top             =   450
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Today is :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   165
         Left            =   45
         TabIndex        =   13
         Top             =   465
         Width           =   630
      End
      Begin VB.Image Image2 
         Height          =   225
         Left            =   0
         Picture         =   "MDIForm1.frx":6CFB
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3330
      End
      Begin VB.Image Image1 
         Height          =   23145
         Index           =   0
         Left            =   3000
         Picture         =   "MDIForm1.frx":6EF3
         Stretch         =   -1  'True
         Top             =   -1560
         Width           =   570
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   990
      Left            =   0
      Picture         =   "MDIForm1.frx":70EB
      ScaleHeight     =   990
      ScaleWidth      =   14685
      TabIndex        =   0
      Top             =   0
      Width           =   14685
      Begin VB.Image Image4 
         Height          =   150
         Left            =   -120
         Picture         =   "MDIForm1.frx":D547
         Top             =   -960
         Width           =   1650
      End
   End
   Begin MSComctlLib.ImageList i32x32 
      Left            =   4800
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":D5FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":E2D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":EFB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":FC8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":10965
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1163F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":12319
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":12FF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":13CCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":149A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":15681
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1635B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":17035
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":17D0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":189E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":196C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A39D
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1B077
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4800
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1BD51
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1CA2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1D705
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1E3DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F0B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1FD93
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":20A6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":21747
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":22421
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":230FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":23DD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":24AAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":25789
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":26463
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2713D
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":27E17
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":28AF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":297CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":29A5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2A109
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2AE63
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2CBFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2E997
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":31E71
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3274B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mLogOut 
         Caption         =   "Log-out"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mRecords 
      Caption         =   "Records"
      Visible         =   0   'False
      Begin VB.Menu mProd 
         Caption         =   "Product Masterlist"
      End
      Begin VB.Menu mPpack 
         Caption         =   "Product Package"
      End
      Begin VB.Menu mPCat 
         Caption         =   "Product Category"
      End
      Begin VB.Menu mMOnitor 
         Caption         =   "Product Monitoring "
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mSI 
         Caption         =   "Sales Invoice"
      End
      Begin VB.Menu mSR 
         Caption         =   "Stock Recieve"
      End
      Begin VB.Menu mPO 
         Caption         =   "Purchase Order"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mSupp 
         Caption         =   "Supplier Masterlist"
      End
   End
   Begin VB.Menu mReport 
      Caption         =   "Reports"
      Visible         =   0   'False
      Begin VB.Menu mProdRep 
         Caption         =   "Products"
      End
      Begin VB.Menu mSupRep 
         Caption         =   "Supplier"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mDSI 
         Caption         =   "Daily Sale Invoice"
      End
      Begin VB.Menu mPr 
         Caption         =   "Product Recieve"
      End
      Begin VB.Menu mPM 
         Caption         =   "Product Monitoring"
      End
   End
   Begin VB.Menu mUtility 
      Caption         =   "Utilities"
      Visible         =   0   'False
      Begin VB.Menu mUA 
         Caption         =   "User Account"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mcalc 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mNote 
         Caption         =   "Notepad"
      End
      Begin VB.Menu mDatetime 
         Caption         =   "Date and Time Settings"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mWE 
         Caption         =   "Windows Explorer"
      End
   End
   Begin VB.Menu mabout 
      Caption         =   "About"
      Visible         =   0   'False
      Begin VB.Menu mconten 
         Caption         =   "Contents..."
      End
      Begin VB.Menu mindex 
         Caption         =   "Index..."
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mAIS 
         Caption         =   "About FV-Inventory System"
      End
   End
   Begin VB.Menu mAction 
      Caption         =   "Action"
      Visible         =   0   'False
      Begin VB.Menu mNew 
         Caption         =   "Create New"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mEdit 
         Caption         =   "Edit Selected"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mDelete 
         Caption         =   "Delete Selected"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mRefresh 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mPrint 
         Caption         =   "Print "
         Shortcut        =   ^P
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mClose 
         Caption         =   "Close"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim bhide As Boolean
Dim getAction As String
 
Private Sub cmdFile_Click()
        
    getAction = "picFile"
    
    If picMenu.Visible = True Then
        Timer1.Interval = 10
        bhide = True
    Else
        Timer1.Interval = 10
        bhide = False
        picMenu.Visible = True
    End If
End Sub

 

Private Sub cmdSet_Click()
    getAction = "picSet"
    
    If picSet.Visible = True Then
        Timer1.Interval = 10
        bhide = True
    Else
        Timer1.Interval = 10
        bhide = False
        picSet.Visible = True
    End If
    
End Sub

 


Private Sub ListView1_Click()
     Dim selItem As ListItem
    On Error GoTo RAE
    Set selItem = Listview1.SelectedItem
    Select Case selItem.Key
        Case "a1": Label3.Caption = "Products"
        Case "a2": Label3.Caption = "Suppliers"
        Case "a3": Label3.Caption = "Category"
        Case "a4": Label3.Caption = "Stock Monitoring"
        Case "a5": Label3.Caption = "Sales Entry"
        Case "b1": Label3.Caption = "Daily Sales"
        Case "b2": Label3.Caption = "Stock Recieve"
    End Select
RAE:
    Set selItem = Nothing
End Sub

Private Sub ListView2_Click()
     Dim selItems As ListItem
    On Error GoTo RAE
    Set selItems = Listview2.SelectedItem
    Select Case selItems.Key
        Case "a6": Label6.Caption = "User Accounts"
        Case "a7": Label6.Caption = "Database Back-up"
        Case "a8": Label6.Caption = "Business Info"
        Case "a9": Label6.Caption = "About System"
    End Select
RAE:
    Set selItems = Nothing
End Sub

Private Sub ListView1_DblClick()
    Dim selItem As ListItem
    
    On Error GoTo RAE
    
    Set selItem = Listview1.SelectedItem
    
    Select Case selItem.Key
        Case "a1":
        Unload frmSuppliers
        Unload frmCategorys
        Unload frmMonitoring
        Unload frmSalesDaily
        Unload frmStockRecieves
        loadForm frmProducts
 
        Case "a2":
        Unload frmCategorys
        Unload frmMonitoring
        Unload frmSalesDaily
        Unload frmStockRecieves
        Unload frmProducts
        loadForm frmSuppliers
        
        Case "a3":
        Unload frmSuppliers
        Unload frmMonitoring
        Unload frmSalesDaily
        Unload frmStockRecieves
        Unload frmProducts
        loadForm frmCategorys
        
        Case "a4":
        Unload frmSuppliers
        Unload frmCategorys
        Unload frmSalesDaily
        Unload frmStockRecieves
        Unload frmProducts
        loadForm frmMonitoring
        
        Case "a5":
        frmSalesInvoice.Show vbModal
        
        Case "b1":
        Unload frmSuppliers
        Unload frmCategorys
        Unload frmMonitoring
        Unload frmStockRecieves
        Unload frmProducts
        loadForm frmSalesDaily
        
        Case "b2":
        Unload frmSuppliers
        Unload frmCategorys
        Unload frmMonitoring
        Unload frmSalesDaily
        Unload frmProducts
        loadForm frmStockRecieves

    End Select
RAE:
    Set selItem = Nothing
End Sub

 

Private Sub Listview2_DblClick()
    Dim List As ListItem
    On Error GoTo RAE
    Set List = Listview2.SelectedItem
    
    Select Case List.Key
        Case "a6"
                frmAEAccount.Show vbModal
        Case "a7"
                frmDBBackup.Show vbModal
        Case "a8":
                frmBusinessInfo.Show vbModal
        Case "a9":
           frmAbout.Show 1
    End Select
RAE:
    Set List = Nothing
End Sub

Private Sub MDIForm_Load()
 frmLogin.Show 1
 Me.Show
 
    Listview1.ListItems.Add , "a1", "Products", 2, 2
    Listview1.ListItems.Add , "a2", "Suppliers", 3, 3
    Listview1.ListItems.Add , "a3", "Category", 6, 6
    Listview1.ListItems.Add , "a4", "Stock Monitoring", 15, 15
    Listview1.ListItems.Add , "a5", "Sales Entry", 16, 16
    Listview1.ListItems.Add , "b1", "Daily Sales", 17, 17
    Listview1.ListItems.Add , "b2", "Stock Recieve", 18, 18
 
    
    Listview2.ListItems.Add , "a6", "User Accounts", 16, 16
    Listview2.ListItems.Add , "a7", "Back-Up", 17, 17
    Listview2.ListItems.Add , "a8", "Business Info", 25, 25
    Listview2.ListItems.Add , "a9", "About ", 19, 19

loadForm frmWelcome
lblCurrentUser = CurrUser.UserNAME
End Sub

 
 
Private Sub MDIForm_Resize()
    On Error Resume Next
    Picture2.Width = MDIForm1.Width
End Sub

Private Sub mWE_Click()
On Error Resume Next
Shell "Explorer.exe", vbNormalFocus
End Sub

 





 
Private Sub mNew_Click()
On Error Resume Next
ActiveForm.Command "New"
End Sub
Private Sub mEdit_Click()
On Error Resume Next
ActiveForm.Command "Edit"
End Sub
Private Sub mRefresh_Click()
On Error Resume Next
ActiveForm.Command "Refresh"
End Sub
Private Sub mDelete_Click()
On Error Resume Next
ActiveForm.Command "Delete"
End Sub
Private Sub mPrint_Click()
On Error Resume Next
ActiveForm.Command "Print"
End Sub
Private Sub mclose_Click()
On Error Resume Next
ActiveForm.Command "Close"
End Sub
 
 
 

Private Sub TreeView1_DblClick()
Select Case TreeView1.SelectedItem.Index
   
    Case 2
            loadForm frmCategorys
    Case 3
            loadForm frmSuppliers
    Case 4
           loadForm frmProducts
     Case 5
            loadForm frmMonitoring
    Case Else
      
End Select
End Sub

 



Private Sub Timer1_Timer()
    Select Case getAction
        Case "picSet"
            If bhide = False Then
                If picSet.Height >= 3495 Then
                    picSet.Height = 3495
                    Timer1.Interval = 0
                    Else
                    picSet.Height = picSet.Height + 249
                End If
                
            Else
            
             If picSet.Height <= 600 Then
                    picSet.Height = 0
                    Timer1.Interval = 0
                    picSet.Visible = False
                    Else
                    picSet.Height = picSet.Height - 249
                    DoEvents
                End If
                
            End If
            
        Case "picFile"
        
            If bhide = False Then
        
                If picMenu.Height >= 4575 Then
                    Else
                    picMenu.Height = picMenu.Height + 249
                End If
        
            Else
             If picMenu.Height <= 600 Then
                    picMenu.Height = 0
                    picMenu.Visible = False
                    Else
                    picMenu.Height = picMenu.Height - 249
                    DoEvents
                End If
            End If
    End Select
End Sub

Private Sub picMenu_Resize()
    cmdSet.Top = picMenu.Top + picMenu.Height
    picSet.Top = picMenu.Top + picMenu.Height + cmdSet.Height
End Sub

Private Sub Timer2_Timer()
lblDate.Caption = Format(Now, "MMM-DD-YYYY- ") & "[ " & Format(Time, "hh:mm:ss am/pm") & " ]"
End Sub
 
