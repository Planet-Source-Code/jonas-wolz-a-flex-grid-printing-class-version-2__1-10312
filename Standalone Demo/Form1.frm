VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "clsTablePrint demo"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows-Standard
   WindowState     =   2  'Maximiert
   Begin VB.Image imgPic 
      Height          =   285
      Left            =   1320
      Picture         =   "Form1.frx":0000
      Top             =   1560
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'For an example how to use this grid, look into
'Form_Click.

'Please be not too annoyed if you discover errors in the English text ;-)
'English is not my native language (and I'm still learning it) !

Private WithEvents mTP As clsTablePrint
Attribute mTP.VB_VarHelpID = -1

Private Sub Form_Click()
    Dim fntSet As IFont, fntSet2 As StdFont
    Dim L As Long, L2 As Long
    
    Me.Cls
    'Set the "start" of the grid
    Me.CurrentY = 150
    
    'With the following code the grid is initialized.
    'If you want to see the effect of the properties simply change the values and see !
    '
    'You'll always have to initialize the following items:
    '- The Cols and Rows properties
    '- The different font properties (ColFont, HeaderFont, FooterFont)
    '- The ColWidth property
    '- Of course the text properties (TextMatrix, HeaderText, FooterText)
    'It's often also a good idea to set...
    '- The CellXOffset and CellYOffset property
    '- The ColAlignment property
    '- The Margin properties (MarginTop, MarginLeft, MarginBottom)
    
    With mTP
        'Set how far the text should keep away from the cells' borders
        .CellXOffset = 60
        .CellYOffset = 15
        .HeaderRows = 2
        'Set the count of cols and rows
        'IMPORTANT: Set those properties first, because all array properties
        ' (e.g. TextMatrix, ColFont, ...) will be reset if you change them !
        .Cols = 5
        .Rows = 6
        'Set the fntSet to the form's Font
        Set fntSet = Me.Font
        'Set the Font for all columns to the form's font.
        'If you pass -1 (or anything <0) to any of the "array" properties of this class
        'all items in the array will be set to the specified value !
        Set .FontMatrix(-1, -1) = Me.Font
        'Clone fntSet (that's why I'm using IFont (fully compatible to StdFont) here ;-)
        fntSet.Clone fntSet2
        'Make the font bold and bigger
        fntSet2.Bold = True
        fntSet2.Size = 10
        'Set the font for the page header
        Set .HeaderFont(-1, -1) = fntSet2
        'like above
        fntSet.Clone fntSet2
        fntSet2.Italic = True
        fntSet2.Name = "Arial"
        fntSet2.Size = 10
        'Set the font for the footer
        Set .FooterFont(-1) = fntSet2
        'Set the last column to a different font !
        fntSet.Clone fntSet2
        fntSet2.Underline = True
        fntSet2.Italic = True
        Set .FontMatrix(-1, .Cols - 1) = fntSet2
        'Set the cell in col 0 in row 3 to another font
        fntSet.Clone fntSet2
        fntSet2.Bold = True
        Set .FontMatrix(2, 0) = fntSet2
        Set fntSet2 = Nothing
        Set fntSet = Nothing
        'Make the line around the header cells a little thicker
        .HeaderLineThickness = 2
        'Show the different alignments
        .ColAlignment(0) = eLeft
        .ColAlignment(1) = eCenter
        .ColAlignment(2) = eRight
        'Set the Top and Left margin:
        .MarginTop = 150
        .MarginLeft = 150
        'If this would be set to false the header would only be printed
        ' on the first page
        .PrintHeaderOnEveryPage = True
                
        'Fill in a example text
        For L = 0 To .Cols - 1
            .HeaderText(-1, L) = "Column " & L
            .FooterText(L) = "Footer " & L
            'TextMatrix is _only_ the text for the cells, _not_ the
            'header or the footer (unlike a FlexGrid)
            For L2 = 0 To .Rows - 1
                .TextMatrix(L2, L) = "Row " & L2 & ", Col " & L
            Next
            'Set the width for the cols:
            .ColWidth(L) = (Me.ScaleWidth - 300) / .Cols
        Next
        'You have to move the text with spaces
        ' (kind of quick and dirty; might be changed in future versions)
        .TextMatrix(1, 0) = "       With Picture !"
        Set .PictureMatrix(1, 0) = imgPic.Picture
        'Demonstrate column merging:
        'Cells are only merged inside a column
        .TextMatrix(3, 2) = .TextMatrix(4, 2)

        
        'Demonstrate Header merging:
        .MergeHeaderRow(0) = True
        .MergeHeaderCol(1) = True
        .HeaderText(0, 4) = .HeaderText(0, 3)
        'Say that merged headers should be centered
        '(sometimes this looks better). If this is False, they'll get
        'the alignment of the *last* merged column.
        .CenterMergedHeader = True
        
        ''-- The following code shows what to do with large pictures: --
        ''Say it should make lines heigh enough for pictures:
        '.ResizeCellsToPicHeight = True
        ''Clear the text:
        '.TextMatrix(5, 1) = ""
        ''And set the picture:
        'Set .PictureMatrix(5, 1) = LoadPicture(App.Path & "\LargePic.bmp")
        
        'Demonstrate Multiline (vbCrLf equals to Chr(13) & Chr(10)):
        .TextMatrix(4, 4) = "These are" & vbCrLf & "multiple lines" & vbCrLf & "of text !"
        
        'Finally draw the Grid on the form:
        ' (Simply change "Me" to "Printer" to print it on the printer !)
        .DrawTable Me
        
        'Tell us how much lines per form:
        Print
        Print "Lines/Form: " & CStr(.CalcNumRowsPerPage(Me))
    End With
End Sub

Private Sub Form_Load()
    Set mTP = New clsTablePrint
    Print "Click on the form to see the demonstration !"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set mTP = Nothing
End Sub



Private Sub mTP_NewPage(objOutput As Object, TopMarginAlreadySet As Boolean, bCancel As Boolean, ByVal lLastPrintedRow As Long)
    'The form is simply cleared in this example here.
    'You should not do this in your programs !
    'If you print the table on a printer, simply call Printer.NewPage in here.
    'If you're making a kind of page preview, you will have to create some
    'sort of multi-page mechanism. (For example, caching all pages as bitmaps (simple but slow) or
    'setting bCancel = True and using lLastPrintedRow + 1 as the lRowToStart parameter to DrawTable()
    'when drawing the next page, etc.)
    
    'Set TopMarginAlreadySet = True if objOutput.CurrentY is the position where
    'the next part of the grid should start. Otherwise the value from MarginTop
    'is added to objOutput.CurrentY.
    
    Me.Cls

End Sub


