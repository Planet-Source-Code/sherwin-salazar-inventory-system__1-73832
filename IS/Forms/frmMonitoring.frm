VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmMonitoring 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   0  'None
   Caption         =   "Monitoring"
   ClientHeight    =   7845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   11160
      TabIndex        =   14
      Top             =   7590
      Width           =   11160
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No Record"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C25418&
         Height          =   255
         Left            =   77
         TabIndex        =   15
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   600
      Width           =   375
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00004080&
      Height          =   375
      Index           =   1
      Left            =   7560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   600
      Width           =   375
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   9360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   11160
      TabIndex        =   0
      Top             =   0
      Width           =   11160
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Monitoring"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   45
         Width           =   3975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Monitoring"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   66
         Width           =   3975
      End
      Begin VB.Image cmdExit 
         Height          =   360
         Left            =   9320
         MouseIcon       =   "frmMonitoring.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmMonitoring.frx":0152
         Stretch         =   -1  'True
         ToolTipText     =   "Close"
         Top             =   -27
         Width           =   360
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6255
      Left            =   30
      TabIndex        =   2
      Top             =   1020
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   11033
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilList"
      SmallIcons      =   "ilList"
      ForeColor       =   -2147483630
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmMonitoring.frx":083C
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product Code"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   5715
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Category"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Supplier"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Quantity"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Quantity Remaining"
         Object.Width           =   3350
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Quantity Sold"
         Object.Width           =   2716
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "UnitPrice"
         Object.Width           =   4480
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Selling Price"
         Object.Width           =   4480
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdButtons 
      Height          =   615
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   360
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1085
      Caption         =   "Reload"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   -2147483642
      cBhover         =   16777215
      LockHover       =   3
      cGradient       =   16119285
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMonitoring.frx":099E
      ImgSize         =   32
      cBack           =   16119285
      mPointer        =   99
      mIcon           =   "frmMonitoring.frx":16F8
   End
   Begin lvButton.lvButtons_H cmdButtons 
      Height          =   615
      Index           =   1
      Left            =   2335
      TabIndex        =   5
      Top             =   360
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   1085
      Caption         =   "Print"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   -2147483642
      cBhover         =   16777215
      LockHover       =   3
      cGradient       =   16119285
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMonitoring.frx":185A
      ImgSize         =   32
      cBack           =   16119285
      mPointer        =   99
      mIcon           =   "frmMonitoring.frx":25B4
   End
   Begin lvButton.lvButtons_H cmdButtons 
      Height          =   615
      Index           =   4
      Left            =   3340
      TabIndex        =   6
      Top             =   360
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1085
      Caption         =   "Close"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   -2147483642
      cBhover         =   16777215
      LockHover       =   3
      cGradient       =   16119285
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMonitoring.frx":2716
      ImgSize         =   32
      cBack           =   16119285
      mPointer        =   99
      mIcon           =   "frmMonitoring.frx":3470
   End
   Begin lvButton.lvButtons_H cmdButtons 
      Height          =   615
      Index           =   2
      Left            =   45
      TabIndex        =   7
      Top             =   360
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1085
      Caption         =   "Adjust"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   -2147483642
      cBhover         =   16777215
      LockHover       =   3
      cGradient       =   16119285
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMonitoring.frx":35D2
      ImgSize         =   32
      cBack           =   16119285
      mPointer        =   99
      mIcon           =   "frmMonitoring.frx":4226
   End
   Begin MSComctlLib.ImageList ilRecordIco 
      Left            =   4680
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoring.frx":4388
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Index           =   1
      Left            =   6435
      TabIndex        =   13
      Top             =   630
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ReOder Stock"
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
      Index           =   1
      Left            =   7995
      TabIndex        =   12
      Top             =   630
      Width           =   1320
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Out of Stock"
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
      Index           =   1
      Left            =   9840
      TabIndex        =   11
      Top             =   630
      Width           =   1200
   End
End
Attribute VB_Name = "frmMonitoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcRecord               As String
Dim srcRecCD                As Variant
Public rsData               As ADODB.Recordset
Dim gtotal                  As Currency
Dim xtotal                  As Currency
Dim xQty                    As Single
Dim xQtySold                As Single
Dim tempx                   As Integer
 

Public Sub LoadEntries()
tempx = 0
ListView1.ListItems.Clear
Set rsData = New ADODB.Recordset
rsData.Open "Select * from tblProduct", CN, adOpenStatic, adLockPessimistic
With rsData
    While Not .EOF
                Dim LS As ListItem
                Set LS = ListView1.ListItems.Add(, , rsData!ProductCode, 1, 1)
                    LS.SubItems(1) = rsData!Description
                    LS.SubItems(2) = rsData!CategoryName
                    LS.SubItems(3) = rsData!SupplierName
                    LS.SubItems(4) = rsData!Qty
                    LS.SubItems(5) = rsData!QtyRemain
                    LS.SubItems(6) = rsData!QtySold
                    LS.SubItems(7) = Format(rsData!UnitPrice, "###,#0.00")
                    LS.SubItems(8) = Format(rsData!SellingPrice, "###,#0.00")
                    
                    xQty = xQty + rsData!Qty
                    xQtySold = xQtySold + rsData!QtySold
                    gtotal = gtotal + Format(rsData!SellingPrice, "###,#0.00")
                    xtotal = xtotal + Format(rsData!UnitPrice, "###,#0.00")
                    
                    tempx = tempx + 1
                    LS.ListSubItems(1).Bold = True
                    LS.ListSubItems(4).Bold = True
                    LS.ListSubItems(5).Bold = True
                    LS.ListSubItems(6).Bold = True
                 
               
                    
             If !QtyRemain = 0 Then
                LS.SubItems(5) = "OUT OF STOCK"
                ListView1.ListItems(tempx).ForeColor = vbRed
                ListView1.ListItems(tempx).ListSubItems(1).ForeColor = vbRed
                ListView1.ListItems(tempx).ListSubItems(2).ForeColor = vbRed
                ListView1.ListItems(tempx).ListSubItems(3).ForeColor = vbRed
                ListView1.ListItems(tempx).ListSubItems(4).ForeColor = vbRed
                ListView1.ListItems(tempx).ListSubItems(5).ForeColor = vbRed
                ListView1.ListItems(tempx).ListSubItems(6).ForeColor = vbRed
                ListView1.ListItems(tempx).ListSubItems(7).ForeColor = vbRed
                ListView1.ListItems(tempx).ListSubItems(8).ForeColor = vbRed
                
            ElseIf !QtyRemain <= 5 Then
                ListView1.ListItems(tempx).ForeColor = &H4080&
                ListView1.ListItems(tempx).ListSubItems(1).ForeColor = &H4080&
                ListView1.ListItems(tempx).ListSubItems(2).ForeColor = &H4080&
                ListView1.ListItems(tempx).ListSubItems(3).ForeColor = &H4080&
                ListView1.ListItems(tempx).ListSubItems(4).ForeColor = &H4080&
                ListView1.ListItems(tempx).ListSubItems(5).ForeColor = &H4080&
                ListView1.ListItems(tempx).ListSubItems(6).ForeColor = &H4080&
                ListView1.ListItems(tempx).ListSubItems(7).ForeColor = &H4080&
                ListView1.ListItems(tempx).ListSubItems(8).ForeColor = &H4080&
           End If
        .MoveNext
         Label1.Caption = "Selected Record: " & ListView1.SelectedItem.Index & "/" & ListView1.ListItems.Count
    Wend
    End With
  tempx = tempx + 1
  Set LS = ListView1.ListItems.Add(, , "")
  LS.SubItems(4) = "Total : " & xQty
  LS.SubItems(6) = "Total : " & xQtySold
  LS.SubItems(7) = "Total : " & Format(xtotal, "###,#0.00")
  LS.SubItems(8) = "Total : " & Format(gtotal, "###,#0.00")
  
  ListView1.ListItems(tempx).ListSubItems(4).ForeColor = vbBlue
  ListView1.ListItems(tempx).ListSubItems(6).ForeColor = vbBlue
  ListView1.ListItems(tempx).ListSubItems(7).ForeColor = vbBlue
  ListView1.ListItems(tempx).ListSubItems(8).ForeColor = vbBlue
  
  LS.ListSubItems(4).Bold = True
  LS.ListSubItems(6).Bold = True
  LS.ListSubItems(7).Bold = True
  LS.ListSubItems(8).Bold = True
End Sub

 

Private Sub cmdButtons_Click(Index As Integer)
Select Case Index

Case 0
On Error Resume Next
Command "Refresh"
Case 1
On Error Resume Next
Command "Print"

Case 2
On Error Resume Next
Command "Edit"

Case 4
On Error Resume Next
Command "Close"

End Select
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call LoadEntries
End Sub

 

Private Sub Form_Unload(Cancel As Integer)
     MDIForm1.mDelete.Enabled = True
     MDIForm1.mEdit.Enabled = True
     MDIForm1.mNew.Enabled = True
     MDIForm1.mPrint.Enabled = True
     loadForm frmWelcome
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    
                
    Case vbKeyF2
               On Error Resume Next
               Command "Edit"
               
    Case vbKeyDelete
               On Error Resume Next
               Command "Delete"
               
    Case vbKeyP Or (KeyCode = 109 And Shift = 2)
               On Error Resume Next
               Command "Print"
               
    Case vbKeyF5
               On Error Resume Next
               Command "Refresh"
               
    Case 67 And Shift = 2
                On Error Resume Next
                Command "Close"
End Select
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then ListView1_Click
End Sub

Private Sub Picture1_Resize()
  cmdExit.Left = Picture1.Width - cmdExit.Width - 23
End Sub


Private Sub Form_Resize()
On Error Resume Next
ListView1.Width = Me.ScaleWidth
ListView1.Height = (Me.ScaleHeight - Picture2.Height) - ListView1.Top
 End Sub

Sub RefreshRecords()
Form_Load
End Sub

 
Private Sub ListView1_Click()
    If Trim(srcRecord) = vbNullString Then
         Label1.Caption = "No Record"
    Else
       Label1.Caption = "Selected Record: " & ListView1.SelectedItem.Index & "/" & ListView1.ListItems.Count
    End If
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
     MDIForm1.mPrint.Visible = True
     MDIForm1.mNew.Enabled = False
     MDIForm1.mDelete.Enabled = False
     MDIForm1.mEdit.Caption = "Adjust Quantity"
     PopupMenu MDIForm1.mAction
End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
SortLV ListView1
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
srcRecCD = ListView1.SelectedItem.Index
srcRecord = ListView1.ListItems.Item(srcRecCD).Text
End Sub

 
Public Sub Command(cmd As String)
Select Case cmd

Case "Edit"
            If Trim(srcRecord) = vbNullString Then
                    MsgBox "Invalid selection.Can't proceed to the operation!", vbExclamation
                    Exit Sub
            Else
           If ListView1.ListItems.Count < 1 Then Exit Sub
           Set rsData = New ADODB.Recordset
           rsData.Open "Select * from tblProduct where ProductCode = '" & ListView1.SelectedItem.Text & "'", CN, adOpenStatic, adLockPessimistic
           If rsData.RecordCount < 1 Then Exit Sub
            With frmAdjustStock
              .Text1.Text = rsData.Fields("ProductCode")
              .Text2.Text = rsData.Fields("Description")
              .Text3.Text = rsData.Fields("QtyRemain")
              .PK = srcRecord
              .Show vbModal
            End With
            End If
Case "Print"
    Set rsData = New ADODB.Recordset
    rsData.Open "tblProduct", CN, adOpenStatic, adLockPessimistic
    Set rptStockMonitoring.DataSource = rsData
    rptStockMonitoring.Show
    
Case "Refresh"
 Me.LoadEntries
 
Case "Close"
      Unload Me


End Select
End Sub




