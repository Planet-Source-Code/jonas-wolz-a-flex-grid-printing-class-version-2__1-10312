Attribute VB_Name = "mdlReadDBGrid"
Option Explicit
'###########################################
'# mdlTPHelper                             #
'# Author: Jonas Wolz                      #
'# This module contains utility            #
'# functions for use with clsTablePrint.   #
'# This module is not needed by the        #
'# class !                                 #
'# --------------------------------------- #
'# Function list:                          #
'# Sub ImportFlexGrid( clsTP As _          #
'#   clsTablePrint, flxGrd As MSFlexGrid): #
'#   This function reads the               #
'#   data from flxGrd into clsTP.          #
'###########################################




'ImportDBGrid:
' This Sub reads the DBGrid specified by dbGrd into clsTP.
' rstData has to be set to the recordset dbGrd gets its data from (it seems to be impossible to get DataSource at runtime !???)
' (e.g. if it's bound to Data1, rstData should be Data1.Recordset)
Sub ImportDBGrid(clsTP As clsTablePrint, dbGrd As DBGrid, rstData As Recordset, Optional ByVal sngDesiredWidth As Single = -1)
    Dim lCol As Long, lRow As Long, rstWork As Recordset
    Dim sngFXGGesWidth As Single, oCol As Column
    
    Set rstWork = rstData.Clone 'Create a copy to work with
    rstData.MoveLast 'So that RecordCount is valid
    rstData.MoveFirst
    clsTP.Rows = rstData.RecordCount
    clsTP.Cols = dbGrd.Columns.Count
    clsTP.HeaderRows = 1
    clsTP.HasFooter = False
    clsTP.LineThickness = 1
    'Use double line width
    clsTP.HeaderLineThickness = 2 * clsTP.LineThickness

    'Set the row height
    clsTP.RowHeightMin = dbGrd.RowHeight
    clsTP.FooterRowHeightMin = dbGrd.RowHeight
    clsTP.HeaderRowHeightMin = dbGrd.RowHeight
    
    'Use some reasonable default values:
    clsTP.CellXOffset = 60
    clsTP.CellYOffset = 30
    clsTP.CenterMergedHeader = False
    clsTP.ResizeCellsToPicHeight = True
    clsTP.PrintHeaderOnEveryPage = True
    
    Set fntOld = New StdFont
    With dbGrd
        sngFXGGesWidth = 0
        Set clsTP.HeaderFont(-1, -1) = .HeadFont
        Set clsTP.FontMatrix(-1, -1) = .Font
        For lCol = 0 To .Columns.Count - 1
            Set oCol = .Columns(lCol)
            Select Case oCol.Alignment
            Case dbgLeft
                clsTP.ColAlignment(lCol) = eLeft
            Case dbgRight
                clsTP.ColAlignment(lCol) = eRight
            Case dbgCenter
                clsTP.ColAlignment(lCol) = eCenter
            Case dbgGeneral
                Select Case rstWork.Fields(oCol.DataField).Type
                Case dbText, dbMemo
                    clsTP.ColAlignment(lCol) = eLeft
                Case Else
                    clsTP.ColAlignment(lCol) = eRight
                End Select
            End Select
            sngFXGGesWidth = sngFXGGesWidth + oCol.Width
            clsTP.HeaderText(0, lCol) = oCol.Caption
        Next
        Do Until rstWork.EOF
            lRow = rstWork.AbsolutePosition
            For lCol = 0 To .Columns.Count - 1
                Set oCol = .Columns(lCol)
                clsTP.TextMatrix(lRow, lCol) = Format(rstWork.Fields(oCol.DataField).Value, oCol.NumberFormat)
                If lRow = 0 Then 'Q&D Hack to save another For...Next loop
                    If sngDesiredWidth > 0 Then
                        clsTP.ColWidth(lCol) = (oCol.Width / sngFXGGesWidth) * sngDesiredWidth
                    Else
                        clsTP.ColWidth(lCol) = oCol.Width
                    End If
                End If
            Next
            rstWork.MoveNext
        Loop
    End With
End Sub




