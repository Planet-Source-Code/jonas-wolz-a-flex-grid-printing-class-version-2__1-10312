VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Demo with DBGrid"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows-Standard
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmDemo.frx":0000
      Height          =   3255
      Left            =   120
      OleObjectBlob   =   "frmDemo.frx":0010
      TabIndex        =   9
      Top             =   360
      Width           =   5295
   End
   Begin VB.Data Data1 
      Caption         =   "Data Source"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programme\Microsoft Visual Basic\Nwind.mdb"
      DefaultCursorType=   0  'Standard-Cursor
      DefaultType     =   2  'ODBC verwenden
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Artikel"
      Top             =   4080
      Width           =   5175
   End
   Begin VB.CheckBox chkColWidth 
      Caption         =   "Resize Col &widths to fill page"
      Height          =   195
      Left            =   6600
      TabIndex        =   8
      Top             =   3840
      Value           =   1  'Aktiviert
      Width           =   2415
   End
   Begin VB.PictureBox picScroll 
      Height          =   3375
      Left            =   5640
      ScaleHeight     =   3315
      ScaleWidth      =   4635
      TabIndex        =   4
      Top             =   360
      Width           =   4695
      Begin VB.VScrollBar vscScroll 
         Height          =   2535
         LargeChange     =   15
         Left            =   4320
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.HScrollBar hscScroll 
         Height          =   255
         LargeChange     =   15
         Left            =   0
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3000
         Width           =   4575
      End
      Begin VB.PictureBox picTarget 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2655
         Left            =   0
         ScaleHeight     =   2625
         ScaleWidth      =   3825
         TabIndex        =   7
         Top             =   0
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh PictureBox"
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print the grid on the printer..."
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "NOTE: Set the DatabaseName AND Recordsource properties of Data1 to reasonable values before running this demo !"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "PictureBox as target:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   1
      Top             =   0
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   5520
      X2              =   5520
      Y1              =   4560
      Y2              =   0
   End
   Begin VB.Label Label1 
      Caption         =   "DBGrid as source:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'The dimensions of the DIN A4 paper size in Twips:
Const A4Height = 16840, A4Width = 11907

'To get the scroll width:
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CYHSCROLL = 3
Private Const SM_CXVSCROLL = 2

'Declared Private WithEvents to get NewPage event:
Private WithEvents cTP As clsTablePrint
Attribute cTP.VB_VarHelpID = -1
Private Sub InitializePictureBox()
    Dim sngVSCWidth As Single, sngHSCHeight As Single
    'Set the size to the DIN A4 width:
    picTarget.Width = A4Width
    picTarget.Height = A4Height
    'Resize the scrollbars:
    sngVSCWidth = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    sngHSCHeight = GetSystemMetrics(SM_CYHSCROLL) * Screen.TwipsPerPixelY
    hscScroll.Move 0, picScroll.ScaleHeight - sngHSCHeight, picScroll.ScaleWidth - sngVSCWidth, sngHSCHeight
    vscScroll.Move picScroll.ScaleWidth - sngVSCWidth, 0, sngVSCWidth, picScroll.ScaleHeight
    
    SetScrollBars
End Sub

Private Sub SetScrollBars()
    hscScroll.Max = (picTarget.Width - picScroll.ScaleWidth + vscScroll.Width) / 120 + 1
    vscScroll.Max = (picTarget.Height - picScroll.ScaleHeight + hscScroll.Height) / 120 + 1
End Sub


Private Sub cmdPrint_Click()
    
    If MsgBox("The application will now print the grid on the default printer (Show a print dialog here later !).", vbInformation + vbOKCancel, "Print") = vbCancel Then Exit Sub
    
    'Simply initialize the printer:
    Printer.Print
    
    'Read the FlexGrid:
    'Set the wanted width of the table to -1 to get the exact widths of the FlexGrid,
    ' to ScaleWidth - [the left and right margins] to get a fitting table !
    ImportDBGrid cTP, DBGrid1, Data1.Recordset, IIf((chkColWidth.Value = vbChecked), Printer.ScaleWidth - 2 * 567, -1)
    
    'Set margins (not needed, but looks better !):
    cTP.MarginBottom = 567 '567 equals to 1 cm
    cTP.MarginLeft = 567
    cTP.MarginTop = 567
    
    'Class begins drawing at CurrentY !
    Printer.CurrentY = cTP.MarginTop
    
    'Finally draw the Grid !
    cTP.DrawTable Printer
    'Done with drawing !
    
    'Say VB it should finally send it:
    Printer.EndDoc
End Sub

Private Sub cmdRefresh_Click()
    
    'Read the DBGrid:
    'Set the wanted width of the table to -1 to get the exact widths of the FlexGrid,
    ' to ScaleWidth - [the left and right margins] to get a fitting table !
    ' You have to pass the recordset with the data to read because the DBGrid only stores
    ' the rows it displays, not the whole recordset -> I need the recordset to read the data.
    ImportDBGrid cTP, DBGrid1, Data1.Recordset, IIf((chkColWidth.Value = vbChecked), picTarget.ScaleWidth - 2 * 567, -1)
    
    'Set margins (not needed, but looks better !):
    cTP.MarginBottom = 567 '567 equals to 1 cm
    cTP.MarginLeft = 567
    cTP.MarginTop = 567
    
    'Clear the box:
    picTarget.Cls
    
    'Class begins drawing at CurrentY !
    picTarget.CurrentY = cTP.MarginTop
    
    'Finally draw the Grid !
    cTP.DrawTable picTarget
    'Done with drawing !
End Sub

Private Sub cTP_NewPage(objOutput As Object, TopMarginAlreadySet As Boolean, bCancel As Boolean, ByVal lLastPrintedRow As Long)
    
    'The class wants a new page, look what to do
    If TypeOf objOutput Is Printer Then
        Printer.NewPage
    Else 'We are printing on the PictureBox !
        objOutput.CurrentY = objOutput.ScaleHeight
        'Simply increase the height of the PicBox here
        ' (very simple, but looks bad in "real" applications)
        objOutput.Height = objOutput.Height + A4Height
        'Draw a line to show the new page:
        objOutput.Line (0, objOutput.CurrentY)-(objOutput.ScaleWidth, objOutput.CurrentY), &H808080
        
        'Set the CurrentY to the position the class should continie with drawing and...
        objOutput.CurrentY = objOutput.CurrentY + cTP.MarginTop
        '... tell it to do so:
        TopMarginAlreadySet = True
        
        'Set the ScrollBar Max properties:
        SetScrollBars
    End If
End Sub

Private Sub Form_Load()
    InitializePictureBox
    Set cTP = New clsTablePrint
    Data1.Refresh
    cmdRefresh_Click
End Sub


Private Sub hscScroll_Change()
    picTarget.Left = -hscScroll.Value * 120
End Sub

Private Sub hscScroll_Scroll()
    hscScroll_Change
End Sub


Private Sub vscScroll_Change()
    picTarget.Top = -vscScroll.Value * 120
End Sub


Private Sub vscScroll_Scroll()
    vscScroll_Change
End Sub


