'--------------------------------------------------------------------------
'
'This code resizes all Columns in a DataGrid
'
'The largest string in a column is measured and
'is taken as the new width of the column
'
'To keep this fast, the code just searches the
'visible plane of the grid. To dynamically change
'the column spacing while browsing, to reflect
'possible changes in width, just call this function
'in your keydown or mouseclick events of the datagrid
'code. Otherwise call it just once after the DataGrid was
'fed with Data.
'
' e.g.:
'
'    ....
'    ....
'    rsTemp.Open cSql, cnTemp, , adOpenStatic, adCmdText
'    Load frmLookUp
'    frmLookUp.Adodc1.ConnectionString = cn.ConnectionString
'    frmLookUp.Adodc1.RecordSource = rsTemp.Source
'    frmLookUp.DataGrid1.MarqueeStyle = dbgHighlightRow
'    frmLookUp.DataGrid1.ClearFields
'    frmLookUp.DataGrid1.ReBind
'    frmLookUp.DataGrid1.Refresh
'    'Whoopiieee, here we go!.......
'    DatagridColumnAutoResize frmLookUp.DataGrid1, frmLookUp
'    ....
'    ....
'
'--------------------------------------------------------------------------




'--------------------------------------------------------------------------
'
' AutoResize all Columns of a Datagrid
'
' Parameters:
'		oDataGrid 	-> the Datagrid to be resized
'		oForm		-> any Form
'
' Note: If ColumnHeading is wider than CellWidth, HeadingWidth is used
'--------------------------------------------------------------------------
Public Sub DatagridColumnAutoResize(ByRef oDataGrid As DataGrid, _
                                    ByRef oForm As Form)
Dim i As Integer, iMax As Integer
Dim t As Integer, tMax As Integer
Dim iWidth As Integer
Dim vBMark As Variant
Dim aWidth As Variant
Dim cText As String
Dim oFont As Font

    On Error Resume Next

    'need this to make TextWidth()
    'work with prossibly different font in DG
    oFont = oForm.Font
    oForm.Font = oDataGrid.Font

    iMax = oDataGrid.Columns.Count - 1
    ReDim aWidth(iMax)

    For i = 0 To iMax   'init maxwidth holder

        aWidth(i) = 0

    Next

    'one visible page to get to an estimate
    tMax = oDataGrid.VisibleRows - 1

    For t = 0 To tMax   'number of rows

        vBMark = oDataGrid.GetBookmark(t)

        For i = 0 To iMax   'number of columns

            cText = oDataGrid.Columns(i).CellText(vBMark)
            iWidth = oForm.TextWidth(cText)

            If iWidth + ((12 * Len(cText)) + 220) > aWidth(i) Then

                'the font is right, the stringlength too, but
                'still some misalignment on long stings. So we
                'have to fiddle this a bit by hand...
                aWidth(i) = iWidth + ((12 * Len(cText)) + 220)

            End If

            If t = 0 Then   'take care of the headers

                iWidth = oForm.TextWidth(oDataGrid.Columns(i).Caption)
                If iWidth + ((12 * Len(cText)) + 220) > aWidth(i) Then

                    aWidth(i) = iWidth + ((12 * Len(cText)) + 220)

                End If

            End If

        Next

    Next

    For i = 0 To iMax   ' finally set the new column width

        oDataGrid.Columns(i).Width = aWidth(i)

    Next

    oForm.Font = oFont

End Sub