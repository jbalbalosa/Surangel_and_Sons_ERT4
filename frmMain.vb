
Imports Microsoft
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Win32

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports System.Windows.Forms.Application
Imports System.Security
Imports System.Security.AccessControl

Public Class frmMain

    Private DirPathEpicor As String = "C:\Users\" & Environment.UserName & "\AppData\Local\Temp\Epicor"
    Private DirPath3Apps As String = "C:\3apps\Temp"
    Private strFileDirectory As String = ""
    Private strExportFileName As String = ""
    Private strExportFullPath As String = ""

    Private strRangeStart As String = ""
    Private strRangeEnd As String = ""
    Private strMultiReport As String = ""

    Dim xlApp As New Office.Interop.Excel.Application
    Dim xlWorkBook As Office.Interop.Excel.Workbook
    Dim xlWorkSheet As Office.Interop.Excel.Worksheet

    Private blnHardwareReport As Boolean = False
    Private blnHardwareReportWCheckBox As Boolean = False
    Private blnFirstReportProcessed As Boolean = False
    Private blnCombineReport As Boolean = False
    Private blnProductExpirationReport As Boolean = False
    Private blnClearOrderPoint As Boolean = False
    Private intCombinedReport As Integer = 0

    Private intDGVPreview As Integer = 0
    Private dgvTemp0 As New DataGridView
    Private dgvTemp1 As New DataGridView
    Private dgvTemp2 As New DataGridView
    Private dgvTemp3 As New DataGridView
    Private dgvTemp4 As New DataGridView
    Private dgvTemp5 As New DataGridView
    Private dgvTemp6 As New DataGridView

    Private strFileName As String = ""
    Private strFilePath As String = ""

    Private strExcelFileNameExpiredProduct As String = ""
    Private strExcelFilePathExpiredProduct As String = ""

    Private strTextFileNameExpiredProduct As String = ""
    Private strTextFilePathExpiredProduct As String = ""

    Private blnHighPerformanceReport As Boolean = False

    Private Function GetSystemColorName(ByVal ColorNumber As Integer) As System.Drawing.Color
        Dim ColorName As System.Drawing.Color

        Select Case ColorNumber
            Case 1 : ColorName = System.Drawing.Color.AliceBlue
            Case 2 : ColorName = System.Drawing.Color.AntiqueWhite
            Case 3 : ColorName = System.Drawing.Color.Aqua
            Case 4 : ColorName = System.Drawing.Color.Aquamarine
            Case 5 : ColorName = System.Drawing.Color.Azure

            Case 6 : ColorName = System.Drawing.Color.Beige
            Case 7 : ColorName = System.Drawing.Color.Bisque
            Case 8 : ColorName = System.Drawing.Color.Black
            Case 9 : ColorName = System.Drawing.Color.BlanchedAlmond
            Case 10 : ColorName = System.Drawing.Color.Blue
            Case 11 : ColorName = System.Drawing.Color.BlueViolet
            Case 12 : ColorName = System.Drawing.Color.Brown
            Case 13 : ColorName = System.Drawing.Color.BurlyWood

            Case 14 : ColorName = System.Drawing.Color.CadetBlue
            Case 15 : ColorName = System.Drawing.Color.Chartreuse
            Case 16 : ColorName = System.Drawing.Color.Chocolate
            Case 17 : ColorName = System.Drawing.Color.Coral
            Case 18 : ColorName = System.Drawing.Color.CornflowerBlue
            Case 19 : ColorName = System.Drawing.Color.Cornsilk
            Case 20 : ColorName = System.Drawing.Color.Crimson
            Case 21 : ColorName = System.Drawing.Color.Cyan

            Case 22 : ColorName = System.Drawing.Color.DarkBlue
            Case 23 : ColorName = System.Drawing.Color.DarkCyan
            Case 24 : ColorName = System.Drawing.Color.DarkGoldenrod
            Case 25 : ColorName = System.Drawing.Color.DarkGray
            Case 26 : ColorName = System.Drawing.Color.DarkGreen
            Case 27 : ColorName = System.Drawing.Color.DarkKhaki
            Case 28 : ColorName = System.Drawing.Color.DarkMagenta
            Case 29 : ColorName = System.Drawing.Color.DarkOliveGreen
            Case 30 : ColorName = System.Drawing.Color.DarkOrange
            Case 31 : ColorName = System.Drawing.Color.DarkOrchid
            Case 32 : ColorName = System.Drawing.Color.DarkRed
            Case 33 : ColorName = System.Drawing.Color.DarkSalmon
            Case 34 : ColorName = System.Drawing.Color.DarkSeaGreen
            Case 35 : ColorName = System.Drawing.Color.DarkSlateBlue
            Case 36 : ColorName = System.Drawing.Color.DarkTurquoise
            Case 37 : ColorName = System.Drawing.Color.DarkViolet
            Case 38 : ColorName = System.Drawing.Color.DeepPink
            Case 39 : ColorName = System.Drawing.Color.DeepSkyBlue
            Case 40 : ColorName = System.Drawing.Color.DimGray
            Case 41 : ColorName = System.Drawing.Color.DodgerBlue

            Case 42 : ColorName = System.Drawing.Color.Firebrick
            Case 43 : ColorName = System.Drawing.Color.FloralWhite
            Case 44 : ColorName = System.Drawing.Color.ForestGreen
            Case 45 : ColorName = System.Drawing.Color.Fuchsia

            Case 46 : ColorName = System.Drawing.Color.Gainsboro
            Case 47 : ColorName = System.Drawing.Color.GhostWhite
            Case 48 : ColorName = System.Drawing.Color.Gold
            Case 49 : ColorName = System.Drawing.Color.Goldenrod
            Case 50 : ColorName = System.Drawing.Color.Gray
            Case 51 : ColorName = System.Drawing.Color.Green
            Case 52 : ColorName = System.Drawing.Color.GreenYellow

            Case 53 : ColorName = System.Drawing.Color.Honeydew
            Case 54 : ColorName = System.Drawing.Color.HotPink
            Case 55 : ColorName = System.Drawing.Color.IndianRed
            Case 56 : ColorName = System.Drawing.Color.Indigo
            Case 57 : ColorName = System.Drawing.Color.Ivory
            Case 58 : ColorName = System.Drawing.Color.Khaki

            Case 59 : ColorName = System.Drawing.Color.Lavender
            Case 60 : ColorName = System.Drawing.Color.LavenderBlush
            Case 61 : ColorName = System.Drawing.Color.LawnGreen
            Case 62 : ColorName = System.Drawing.Color.LemonChiffon
            Case 63 : ColorName = System.Drawing.Color.LightBlue
            Case 64 : ColorName = System.Drawing.Color.LightCoral
            Case 65 : ColorName = System.Drawing.Color.LightCyan
            Case 66 : ColorName = System.Drawing.Color.LightGoldenrodYellow
            Case 67 : ColorName = System.Drawing.Color.LightGray
            Case 68 : ColorName = System.Drawing.Color.LightGreen
            Case 69 : ColorName = System.Drawing.Color.LightPink
            Case 70 : ColorName = System.Drawing.Color.LightSalmon
            Case 71 : ColorName = System.Drawing.Color.LightSeaGreen
            Case 72 : ColorName = System.Drawing.Color.LightSkyBlue
            Case 73 : ColorName = System.Drawing.Color.LightSlateGray
            Case 74 : ColorName = System.Drawing.Color.LightSteelBlue
            Case 75 : ColorName = System.Drawing.Color.LightYellow
            Case 76 : ColorName = System.Drawing.Color.Lime
            Case 77 : ColorName = System.Drawing.Color.LimeGreen
            Case 78 : ColorName = System.Drawing.Color.Linen

            Case 79 : ColorName = System.Drawing.Color.Magenta
            Case 80 : ColorName = System.Drawing.Color.Maroon
            Case 81 : ColorName = System.Drawing.Color.MediumAquamarine
            Case 82 : ColorName = System.Drawing.Color.MediumBlue
            Case 83 : ColorName = System.Drawing.Color.MediumOrchid
            Case 84 : ColorName = System.Drawing.Color.MediumPurple
            Case 85 : ColorName = System.Drawing.Color.MediumSeaGreen
            Case 86 : ColorName = System.Drawing.Color.MediumSlateBlue
            Case 87 : ColorName = System.Drawing.Color.MediumSpringGreen
            Case 88 : ColorName = System.Drawing.Color.MediumTurquoise
            Case 89 : ColorName = System.Drawing.Color.MediumVioletRed
            Case 90 : ColorName = System.Drawing.Color.MidnightBlue
            Case 91 : ColorName = System.Drawing.Color.MintCream
            Case 92 : ColorName = System.Drawing.Color.MistyRose
            Case 93 : ColorName = System.Drawing.Color.Moccasin

            Case 94 : ColorName = System.Drawing.Color.NavajoWhite
            Case 95 : ColorName = System.Drawing.Color.Navy

            Case 96 : ColorName = System.Drawing.Color.OldLace
            Case 97 : ColorName = System.Drawing.Color.Olive
            Case 98 : ColorName = System.Drawing.Color.OliveDrab
            Case 99 : ColorName = System.Drawing.Color.Orange
            Case 100 : ColorName = System.Drawing.Color.OrangeRed
            Case 101 : ColorName = System.Drawing.Color.Orchid

            Case 102 : ColorName = System.Drawing.Color.PaleGreen
            Case 103 : ColorName = System.Drawing.Color.PaleTurquoise
            Case 104 : ColorName = System.Drawing.Color.PaleVioletRed
            Case 105 : ColorName = System.Drawing.Color.PapayaWhip
            Case 106 : ColorName = System.Drawing.Color.PeachPuff
            Case 107 : ColorName = System.Drawing.Color.Peru
            Case 108 : ColorName = System.Drawing.Color.Pink
            Case 109 : ColorName = System.Drawing.Color.Plum
            Case 110 : ColorName = System.Drawing.Color.PowderBlue
            Case 111 : ColorName = System.Drawing.Color.Purple

            Case 112 : ColorName = System.Drawing.Color.Red
            Case 113 : ColorName = System.Drawing.Color.RosyBrown
            Case 114 : ColorName = System.Drawing.Color.RoyalBlue

            Case 115 : ColorName = System.Drawing.Color.SkyBlue
            Case 116 : ColorName = System.Drawing.Color.SlateBlue
            Case 117 : ColorName = System.Drawing.Color.SlateGray
            Case 118 : ColorName = System.Drawing.Color.Snow
            Case 119 : ColorName = System.Drawing.Color.SpringGreen
            Case 120 : ColorName = System.Drawing.Color.SteelBlue

            Case 121 : ColorName = System.Drawing.Color.Tan
            Case 122 : ColorName = System.Drawing.Color.Teal
            Case 123 : ColorName = System.Drawing.Color.Thistle
            Case 124 : ColorName = System.Drawing.Color.Tomato
            Case 125 : ColorName = System.Drawing.Color.Transparent
            Case 126 : ColorName = System.Drawing.Color.Turquoise

            Case 127 : ColorName = System.Drawing.Color.Violet

            Case 128 : ColorName = System.Drawing.Color.Wheat
            Case 129 : ColorName = System.Drawing.Color.White
            Case 130 : ColorName = System.Drawing.Color.WhiteSmoke

            Case 131 : ColorName = System.Drawing.Color.Yellow
            Case 132 : ColorName = System.Drawing.Color.YellowGreen

        End Select

        Return ColorName

    End Function

    Private Function GetSystemColorNumber(ByVal ColorName As System.Drawing.Color) As Integer
        Dim ColorNumber As Integer

        Select Case ColorName
            Case System.Drawing.Color.AliceBlue : ColorNumber = 1
            Case System.Drawing.Color.AntiqueWhite : ColorNumber = 2
            Case System.Drawing.Color.Aqua : ColorNumber = 3
            Case System.Drawing.Color.Aquamarine : ColorNumber = 4
            Case System.Drawing.Color.Azure : ColorNumber = 5

            Case System.Drawing.Color.Beige : ColorNumber = 6
            Case System.Drawing.Color.Bisque : ColorNumber = 7
            Case System.Drawing.Color.Black : ColorNumber = 8
            Case System.Drawing.Color.BlanchedAlmond : ColorNumber = 9
            Case System.Drawing.Color.Blue : ColorNumber = 10
            Case System.Drawing.Color.BlueViolet : ColorNumber = 11
            Case System.Drawing.Color.Brown : ColorNumber = 12
            Case System.Drawing.Color.BurlyWood : ColorNumber = 13

            Case System.Drawing.Color.CadetBlue : ColorNumber = 14
            Case System.Drawing.Color.Chartreuse : ColorNumber = 15
            Case System.Drawing.Color.Chocolate : ColorNumber = 16
            Case System.Drawing.Color.Coral : ColorNumber = 17
            Case System.Drawing.Color.CornflowerBlue : ColorNumber = 18
            Case System.Drawing.Color.Cornsilk : ColorNumber = 19
            Case System.Drawing.Color.Crimson : ColorNumber = 20
            Case System.Drawing.Color.Cyan : ColorNumber = 21

            Case System.Drawing.Color.DarkBlue : ColorNumber = 22
            Case System.Drawing.Color.DarkCyan : ColorNumber = 23
            Case System.Drawing.Color.DarkGoldenrod : ColorNumber = 24
            Case System.Drawing.Color.DarkGray : ColorNumber = 25
            Case System.Drawing.Color.DarkGreen : ColorNumber = 26
            Case System.Drawing.Color.DarkKhaki : ColorNumber = 27
            Case System.Drawing.Color.DarkMagenta : ColorNumber = 28
            Case System.Drawing.Color.DarkOliveGreen : ColorNumber = 29
            Case System.Drawing.Color.DarkOrange : ColorNumber = 30
            Case System.Drawing.Color.DarkOrchid : ColorNumber = 31
            Case System.Drawing.Color.DarkRed : ColorNumber = 32
            Case System.Drawing.Color.DarkSalmon : ColorNumber = 33
            Case System.Drawing.Color.DarkSeaGreen : ColorNumber = 34
            Case System.Drawing.Color.DarkSlateBlue : ColorNumber = 35
            Case System.Drawing.Color.DarkTurquoise : ColorNumber = 36
            Case System.Drawing.Color.DarkViolet : ColorNumber = 37
            Case System.Drawing.Color.DeepPink : ColorNumber = 38
            Case System.Drawing.Color.DeepSkyBlue : ColorNumber = 39
            Case System.Drawing.Color.DimGray : ColorNumber = 40
            Case System.Drawing.Color.DodgerBlue : ColorNumber = 40

            Case System.Drawing.Color.Firebrick : ColorNumber = 42
            Case System.Drawing.Color.FloralWhite : ColorNumber = 43
            Case System.Drawing.Color.ForestGreen : ColorNumber = 44
            Case System.Drawing.Color.Fuchsia : ColorNumber = 45

            Case System.Drawing.Color.Gainsboro : ColorNumber = 46
            Case System.Drawing.Color.GhostWhite : ColorNumber = 47
            Case System.Drawing.Color.Gold : ColorNumber = 48
            Case System.Drawing.Color.Goldenrod : ColorNumber = 49
            Case System.Drawing.Color.Gray : ColorNumber = 50
            Case System.Drawing.Color.Green : ColorNumber = 51
            Case System.Drawing.Color.GreenYellow : ColorNumber = 52

            Case System.Drawing.Color.Honeydew : ColorNumber = 53
            Case System.Drawing.Color.HotPink : ColorNumber = 54
            Case System.Drawing.Color.IndianRed : ColorNumber = 55
            Case System.Drawing.Color.Indigo : ColorNumber = 56
            Case System.Drawing.Color.Ivory : ColorNumber = 57

            Case System.Drawing.Color.Khaki : ColorNumber = 58

            Case System.Drawing.Color.Lavender : ColorNumber = 59
            Case System.Drawing.Color.LavenderBlush : ColorNumber = 60
            Case System.Drawing.Color.LawnGreen : ColorNumber = 61
            Case System.Drawing.Color.LemonChiffon : ColorNumber = 62
            Case System.Drawing.Color.LightBlue : ColorNumber = 63
            Case System.Drawing.Color.LightCoral : ColorNumber = 64
            Case System.Drawing.Color.LightCyan : ColorNumber = 65
            Case System.Drawing.Color.LightGoldenrodYellow : ColorNumber = 66
            Case System.Drawing.Color.LightGray : ColorNumber = 67
            Case System.Drawing.Color.LightGreen : ColorNumber = 68
            Case System.Drawing.Color.LightPink : ColorNumber = 69
            Case System.Drawing.Color.LightSalmon : ColorNumber = 70
            Case System.Drawing.Color.LightSeaGreen : ColorNumber = 71
            Case System.Drawing.Color.LightSkyBlue : ColorNumber = 72
            Case System.Drawing.Color.LightSlateGray : ColorNumber = 73
            Case System.Drawing.Color.LightSteelBlue : ColorNumber = 74
            Case System.Drawing.Color.LightYellow : ColorNumber = 75
            Case System.Drawing.Color.Lime : ColorNumber = 76
            Case System.Drawing.Color.LimeGreen : ColorNumber = 77
            Case System.Drawing.Color.Linen : ColorNumber = 78

            Case System.Drawing.Color.Magenta : ColorNumber = 79
            Case System.Drawing.Color.Maroon : ColorNumber = 80
            Case System.Drawing.Color.MediumAquamarine : ColorNumber = 81
            Case System.Drawing.Color.MediumBlue : ColorNumber = 82
            Case System.Drawing.Color.MediumOrchid : ColorNumber = 83
            Case System.Drawing.Color.MediumPurple : ColorNumber = 84
            Case System.Drawing.Color.MediumSeaGreen : ColorNumber = 85
            Case System.Drawing.Color.MediumSlateBlue : ColorNumber = 86
            Case System.Drawing.Color.MediumSpringGreen : ColorNumber = 87
            Case System.Drawing.Color.MediumTurquoise : ColorNumber = 88
            Case System.Drawing.Color.MediumVioletRed : ColorNumber = 89
            Case System.Drawing.Color.MidnightBlue : ColorNumber = 90
            Case System.Drawing.Color.MintCream : ColorNumber = 91
            Case System.Drawing.Color.MistyRose : ColorNumber = 92
            Case System.Drawing.Color.Moccasin : ColorNumber = 93

            Case System.Drawing.Color.NavajoWhite : ColorNumber = 94
            Case System.Drawing.Color.Navy : ColorNumber = 95

            Case System.Drawing.Color.OldLace : ColorNumber = 96
            Case System.Drawing.Color.Olive : ColorNumber = 97
            Case System.Drawing.Color.OliveDrab : ColorNumber = 98
            Case System.Drawing.Color.Orange : ColorNumber = 99
            Case System.Drawing.Color.OrangeRed : ColorNumber = 100
            Case System.Drawing.Color.Orchid : ColorNumber = 101

            Case System.Drawing.Color.PaleGreen : ColorNumber = 102
            Case System.Drawing.Color.PaleTurquoise : ColorNumber = 103
            Case System.Drawing.Color.PaleVioletRed : ColorNumber = 104
            Case System.Drawing.Color.PapayaWhip : ColorNumber = 105
            Case System.Drawing.Color.PeachPuff : ColorNumber = 106
            Case System.Drawing.Color.Peru : ColorNumber = 107
            Case System.Drawing.Color.Pink : ColorNumber = 108
            Case System.Drawing.Color.Plum : ColorNumber = 109
            Case System.Drawing.Color.PowderBlue : ColorNumber = 110
            Case System.Drawing.Color.Purple : ColorNumber = 111

            Case System.Drawing.Color.Red : ColorNumber = 112
            Case System.Drawing.Color.RosyBrown : ColorNumber = 113
            Case System.Drawing.Color.RoyalBlue : ColorNumber = 114

            Case System.Drawing.Color.SkyBlue : ColorNumber = 115
            Case System.Drawing.Color.SlateBlue : ColorNumber = 116
            Case System.Drawing.Color.SlateGray : ColorNumber = 117
            Case System.Drawing.Color.Snow : ColorNumber = 118
            Case System.Drawing.Color.SpringGreen : ColorNumber = 119
            Case System.Drawing.Color.SteelBlue : ColorNumber = 120

            Case System.Drawing.Color.Tan : ColorNumber = 121
            Case System.Drawing.Color.Teal : ColorNumber = 122
            Case System.Drawing.Color.Thistle : ColorNumber = 1123
            Case System.Drawing.Color.Tomato : ColorNumber = 124
            Case System.Drawing.Color.Transparent : ColorNumber = 125
            Case System.Drawing.Color.Turquoise : ColorNumber = 126

            Case System.Drawing.Color.Violet : ColorNumber = 127

            Case System.Drawing.Color.Wheat : ColorNumber = 128
            Case System.Drawing.Color.White : ColorNumber = 129
            Case System.Drawing.Color.WhiteSmoke : ColorNumber = 130

            Case System.Drawing.Color.Yellow : ColorNumber = 131
            Case System.Drawing.Color.YellowGreen : ColorNumber = 132

        End Select

        Return ColorNumber

    End Function


    Private Sub CheckHardwareReportFormat()
        blnHardwareReport = False
        blnHardwareReportWCheckBox = False

        With dgvPreviewExcel
            If .RowCount <> 0 Then
                If LCase(Trim(.Columns(0).HeaderText)) = "" And LCase(Trim(.Columns(1).HeaderText)) = "st" And LCase(Trim(.Columns(2).HeaderText)) = "sku" And LCase(Trim(.Columns(3).HeaderText)) = "description" And LCase(Trim(.Columns(4).HeaderText)) = "upc" Then
                    blnHardwareReport = True
                    blnHardwareReportWCheckBox = True
                ElseIf (LCase(Trim(.Columns(0).HeaderText)) = "st" And LCase(Trim(.Columns(1).HeaderText)) = "sku" And LCase(Trim(.Columns(2).HeaderText)) = "description" And LCase(Trim(.Columns(3).HeaderText)) = "upc") Then
                    blnHardwareReport = True
                End If
            End If
        End With
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub ExportDiscontinuedItem(ByVal SourceWorkSheet As Worksheet, ByVal WorkSheet As Worksheet)
        Dim intLoop As Integer = 2
        Dim intDiscountinue As Integer = 2

        Me.Cursor = Cursors.WaitCursor
        With WorkSheet
            .Cells(1, 1) = "Store"
            .Cells(1, 2) = "Class"
            .Cells(1, 3) = "SKU"
            .Cells(1, 4) = "Qty Avail"
            .Cells(1, 5) = "Description"
            .Cells(1, 6) = "Date Last Sale"
            .Cells(1, 7) = "Date Last Receipt"
            .Cells(1, 8) = "Last Physical"
            .Cells(1, 9) = "Pop Code"
            .Cells(1, 10) = "Prime Vendor"
            .Cells(1, 11) = "Alt Vendor"
            .Cells(1, 12) = "Retail Price"
            .Cells(1, 13) = "Location"
            .Cells(1, 14) = "UPC"
            .Cells(1, 15) = "Discontinue"
            .Cells(1, 16) = "Store Closeout"
            .Cells(1, 17) = "Act GP%"
            .Cells(1, 18) = "Overstock"

            With .Range("A1", "R1")
                .EntireRow.WrapText = True
                .EntireRow.VerticalAlignment = XlVAlign.xlVAlignCenter
                .Font.Bold = True
                .Cells.Interior.Color = RGB(173, 208, 239)
                .EntireRow.HorizontalAlignment = XlHAlign.xlHAlignCenter
            End With

            lblProcessing.Text = "Status: Exporting Discontinued Items..."
            Do While SourceWorkSheet.Range("A" & intLoop).Value <> ""
                DoEvents()
                If SourceWorkSheet.Range("AK" & intLoop).Value = "Y" Then
                    .Cells(intDiscountinue, 1) = "'" & SourceWorkSheet.Range("A" & intLoop).Value
                    .Cells(intDiscountinue, 2) = SourceWorkSheet.Range("AD" & intLoop).Value
                    .Cells(intDiscountinue, 3) = SourceWorkSheet.Range("B" & intLoop).Value
                    .Cells(intDiscountinue, 4) = SourceWorkSheet.Range("I" & intLoop).Value
                    .Cells(intDiscountinue, 5) = SourceWorkSheet.Range("C" & intLoop).Value
                    .Cells(intDiscountinue, 6) = SourceWorkSheet.Range("Z" & intLoop).Value '"Date Last Sale"
                    .Cells(intDiscountinue, 7) = SourceWorkSheet.Range("AA" & intLoop).Value '"Date Last Receipt"
                    .Cells(intDiscountinue, 8) = SourceWorkSheet.Range("AF" & intLoop).Value '"Last Physical"
                    .Cells(intDiscountinue, 9) = SourceWorkSheet.Range("AG" & intLoop).Value '"Pop Code"
                    .Cells(intDiscountinue, 10) = SourceWorkSheet.Range("AC" & intLoop).Value '"Prime Vendor"
                    .Cells(intDiscountinue, 11) = SourceWorkSheet.Range("AD" & intLoop).Value '"Alt Vendor"
                    .Cells(intDiscountinue, 12) = SourceWorkSheet.Range("Y" & intLoop).Value '"Retail Price"
                    .Cells(intDiscountinue, 13) = SourceWorkSheet.Range("AH" & intLoop).Value '"Location"
                    .Cells(intDiscountinue, 14) = SourceWorkSheet.Range("AI" & intLoop).Value '"UPC"
                    .Cells(intDiscountinue, 15) = SourceWorkSheet.Range("AK" & intLoop).Value '"Discontinue"
                    .Cells(intDiscountinue, 16) = SourceWorkSheet.Range("AV" & intLoop).Value '"Store Closeout"
                    .Cells(intDiscountinue, 17) = SourceWorkSheet.Range("AN" & intLoop).Value '"Act GP%"
                    '.Cells(intDiscountinue, 18) = SourceWorkSheet.Range("B" & intLoop).Value '"Overstock"

                    lblProcessing.Text = "Status: Exporting Discontinued Items..." & intDiscountinue
                    intDiscountinue += 1

                End If
                intLoop += 1
            Loop

            With .Range("A1", "R" & intDiscountinue + 1)
                .EntireColumn.AutoFit()
                .Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                .Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                .Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                .Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                .Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
                .Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
            End With

            .Range("A1", "R1").Select()
            xlApp.ActiveWindow.SplitColumn = 3
            xlApp.ActiveWindow.SplitRow = 1
            xlApp.ActiveWindow.FreezePanes = True

            With .PageSetup
                .BottomMargin = 22
                .CenterFooter = "&P"
                .CenterHorizontally = True
                .FooterMargin = 11
                .LeftMargin = 0
                .Orientation = Orientation(cboPaperOrientation.Text)
                .PaperSize = PaperSize(cboPaperSize.Text)
                .PrintArea = "A1:" & "R" & intDiscountinue + 1
                .PrintTitleRows = "$1:$1"
                .PrintTitleColumns = "$A" & ":$R"
                .RightFooter = "&D&T"
                .RightMargin = 0
                .TopMargin = 0
                .Zoom = 100
            End With
        End With

        Me.Cursor = Cursors.Default
        lblProcessing.Text = "Status: Done."

    End Sub

    Private Sub ComputeAvailableQuantity()
        Dim intRow As Integer

        lblProcessing.Text = "Status: Calculating Available Quantity..."
        Me.Cursor = Cursors.WaitCursor
        For intRow = 0 To dgvPreviewExcel.RowCount - 1
            DoEvents()
            lblProcessing.Text = "Status: Calculating Available Quantity... Row " & intRow & " of " & dgvPreviewExcel.RowCount
            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(6).Value) = True Then
                dgvPreviewExcel.Rows(intRow).Cells(6).Value = "0.00"
            End If
            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(7).Value) = True Then
                dgvPreviewExcel.Rows(intRow).Cells(7).Value = "0.00"
            End If

            dgvPreviewExcel.Rows(intRow).Cells(8).Value = Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(6).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(7).Value)
        Next
        Me.Cursor = Cursors.Default
        lblProcessing.Text = "Status: The data has been loaded and ready to process."

    End Sub

    Private Sub ExportDataToExcel(ByVal DataGridView As DataGridView)
        'Try

        Me.Cursor = Cursors.WaitCursor
        Dim strFileName As String = Path.GetFileNameWithoutExtension(strExportFullPath)
        Dim strDateTime As String = FormatDateTime(Now, DateFormat.ShortDate) & FormatDateTime(Now, DateFormat.ShortTime)
        Dim strVendorLookupName As String = ""
        Dim intRemoveRow As Integer = 0

        Me.Cursor = Cursors.WaitCursor
        strDateTime = Replace(strDateTime, "/", "")
        strDateTime = Replace(strDateTime, ":", "")
        Me.Cursor = Cursors.WaitCursor

        If xlApp Is Nothing Then
            MessageBox.Show("There's no Microsoft Excel installed not found")
            Exit Sub
        End If

        xlWorkBook = CType(xlApp.Workbooks.Add(), Office.Interop.Excel.Workbook)
        xlWorkSheet = xlWorkBook.Worksheets(1)
        xlWorkSheet.Name = "Z Report"

        tpbExport.Maximum = dgvPreviewExcel.Rows.Count
        tpbExport.Visible = True
        'dgvPreviewExcel.AllowUserToAddRows = False


        Dim intColumn As Integer = 0
        Dim intRow As Integer = 0

        Dim intBaseRow As Integer = 0
        Dim intExcelSheetRow As Integer = 2

        Dim dblQOH As Double = 0

        'For intRow = 0 To dgvPreviewExcel.Rows.Count - 1
        '    DoEvents()

        '    lblProcessing.Text = "Status: Processing Row #:" & intRow + 1
        '    tpbExport.Value = intRow + 1
        '    If dgvPreviewExcel.Rows(intRow).Cells(6).Value.ToString = "" Then
        '        dgvPreviewExcel.Rows(intRow).Cells(6).Value = 0
        '    End If

        '    If dgvPreviewExcel.Rows(intRow).Cells(7).Value.ToString = "" Then
        '        dgvPreviewExcel.Rows(intRow).Cells(7).Value = 0
        '    End If
        '    dgvPreviewExcel.Rows(intRow).Cells(8).Value = dgvPreviewExcel.Rows(intRow).Cells(6).Value + dgvPreviewExcel.Rows(intRow).Cells(7).Value

        'Next

        For intRow = 0 To dgvPreviewExcel.Rows.Count - 1
            DoEvents()
            lblProcessing.Text = "Status: Processing Row #:" & intRow


            If blnHighPerformanceReport = False Then

                If strVendorLookupName = "" Then
                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells("Vendor").Value) = False Then
                        strVendorLookupName = dgvPreviewExcel.Rows(intRow).Cells("Vendor").Value
                    End If
                End If

            End If

            If intRow = 0 Then
                If dgvPreviewExcel.Rows(intRow).Cells(0).Value.ToString <> "" Then
                    If dgvPreviewExcel.Rows(intRow).Cells(0).Value <> 6 Then
                        For intColumn = 0 To dgvPreviewExcel.Columns.Count - 1
                            DoEvents()
                            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(intColumn).Value) = False Then
                                xlWorkSheet.Cells(intExcelSheetRow, intColumn + 1) = CharPrior(intColumn) & Trim(dgvPreviewExcel.Rows(intRow).Cells(intColumn).Value)
                            End If
                        Next
                        intBaseRow += 1
                    End If
                End If

                If chkClearOrderPoint.Checked = True Then
                    xlWorkSheet.Cells(intExcelSheetRow, 10) = ""
                End If

            ElseIf intRow >= 1 Then
                If dgvPreviewExcel.Rows(intRow).Cells(6).Value.ToString = "" Then
                    dgvPreviewExcel.Rows(intRow).Cells(6).Value = 0
                End If

                If dgvPreviewExcel.Rows(intRow).Cells(7).Value.ToString = "" Then
                    dgvPreviewExcel.Rows(intRow).Cells(7).Value = 0
                End If
                dgvPreviewExcel.Rows(intRow).Cells(8).Value = dgvPreviewExcel.Rows(intRow).Cells(6).Value + dgvPreviewExcel.Rows(intRow).Cells(7).Value


                If dgvPreviewExcel.Rows(intRow).Cells(0).Value <> 6 Then
                    If dgvPreviewExcel.Rows(intRow).Cells(1).Value.ToString = dgvPreviewExcel.Rows(intRow - 1).Cells(1).Value.ToString Then

                        If dgvPreviewExcel.Rows(intRow).Cells(6).Value.ToString <> "" Then

                            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(6).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(0).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(0).Value.ToString Then
                                'xlWorkSheet.Cells(intExcelSheetRow, 7) = dgvPreviewExcel.Rows(intRow).Cells(6).Value
                                xlWorkSheet.Cells(intExcelSheetRow, 7) = xlWorkSheet.Range("G" & intExcelSheetRow).Value + dgvPreviewExcel.Rows(intRow).Cells(6).Value   '... removed to prevent adding of QOH with the same item number
                            End If

                            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(7).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(0).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(0).Value.ToString Then
                                xlWorkSheet.Cells(intExcelSheetRow, 8) = xlWorkSheet.Range("H" & intExcelSheetRow).Value + dgvPreviewExcel.Rows(intRow).Cells(7).Value
                                'xlWorkSheet.Cells(intExcelSheetRow, 8) = dgvPreviewExcel.Rows(intRow).Cells(7).Value
                            End If

                            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(8).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(0).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(0).Value.ToString Then
                                xlWorkSheet.Cells(intExcelSheetRow, 9) = xlWorkSheet.Range("I" & intExcelSheetRow).Value + dgvPreviewExcel.Rows(intRow).Cells(8).Value
                                'xlWorkSheet.Cells(intExcelSheetRow, 9) = dgvPreviewExcel.Rows(intRow).Cells(8).Value
                            End If

                            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(9).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(0).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(0).Value.ToString Then
                                xlWorkSheet.Cells(intExcelSheetRow, 10) = "" 'xlWorkSheet.Range("J" & intExcelSheetRow).Value + dgvPreviewExcel.Rows(intRow).Cells(9).Value
                                'xlWorkSheet.Cells(intExcelSheetRow, 10) = dgvPreviewExcel.Rows(intRow).Cells(9).Value
                            End If

                            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(10).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(0).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(0).Value.ToString Then
                                If blnHighPerformanceReport = False Then
                                    xlWorkSheet.Cells(intExcelSheetRow, 11) = xlWorkSheet.Range("K" & intExcelSheetRow).Value + dgvPreviewExcel.Rows(intRow).Cells(10).Value
                                Else
                                    xlWorkSheet.Cells(intExcelSheetRow, 11) = xlWorkSheet.Range("K" & intExcelSheetRow).Value
                                End If
                                'xlWorkSheet.Cells(intExcelSheetRow, 11) = dgvPreviewExcel.Rows(intRow).Cells(10).Value
                            End If

                            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(11).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(0).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(0).Value.ToString Then
                                If blnHighPerformanceReport = False Then
                                    xlWorkSheet.Cells(intExcelSheetRow, 12) = xlWorkSheet.Range("L" & intExcelSheetRow).Value + dgvPreviewExcel.Rows(intRow).Cells(11).Value
                                Else
                                    xlWorkSheet.Cells(intExcelSheetRow, 12) = xlWorkSheet.Range("L" & intExcelSheetRow).Value
                                End If
                                'xlWorkSheet.Cells(intExcelSheetRow, 12) = dgvPreviewExcel.Rows(intRow).Cells(11).Value
                            End If

                            If dgvPreviewExcel.Rows(intRow).Cells(12).Value.ToString <> "" Then
                                If dgvPreviewExcel.Rows(intRow).Cells(0).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(0).Value.ToString Then

                                    'If blnHighPerformanceReport = False Then
                                    xlWorkSheet.Cells(intExcelSheetRow, 13) = xlWorkSheet.Range("M" & intExcelSheetRow).Value + dgvPreviewExcel.Rows(intRow).Cells(12).Value
                                    'Else
                                    '    xlWorkSheet.Cells(intExcelSheetRow, 13) = xlWorkSheet.Range("M" & intExcelSheetRow).Value
                                    'End If
                                    'xlWorkSheet.Cells(intExcelSheetRow, 13) = dgvPreviewExcel.Rows(intRow).Cells(12).Value
                                End If
                            End If

                            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(14).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(0).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(0).Value.ToString Then
                                xlWorkSheet.Cells(intExcelSheetRow, 15) = xlWorkSheet.Range("O" & intExcelSheetRow).Value + dgvPreviewExcel.Rows(intRow).Cells(14).Value
                                'xlWorkSheet.Cells(intExcelSheetRow, 15) = dgvPreviewExcel.Rows(intRow).Cells(14).Value
                            End If

                            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(15).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(0).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(0).Value.ToString Then
                                xlWorkSheet.Cells(intExcelSheetRow, 16) = xlWorkSheet.Range("P" & intExcelSheetRow).Value '+ dgvPreviewExcel.Rows(intRow).Cells(15).Value
                                'xlWorkSheet.Cells(intExcelSheetRow, 16) = dgvPreviewExcel.Rows(intRow).Cells(15).Value
                            End If

                            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(16).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(0).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(0).Value.ToString Then
                                xlWorkSheet.Cells(intExcelSheetRow, 17) = xlWorkSheet.Range("Q" & intExcelSheetRow).Value + dgvPreviewExcel.Rows(intRow).Cells(16).Value
                                'xlWorkSheet.Cells(intExcelSheetRow, 17) = dgvPreviewExcel.Rows(intRow).Cells(16).Value
                            End If

                            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(18).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(0).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(0).Value.ToString Then
                                xlWorkSheet.Cells(intExcelSheetRow, 19) = xlWorkSheet.Range("S" & intExcelSheetRow).Value + dgvPreviewExcel.Rows(intRow).Cells(18).Value
                                'xlWorkSheet.Cells(intExcelSheetRow, 19) = dgvPreviewExcel.Rows(intRow).Cells(18).Value
                            End If

                            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(19).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(0).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(0).Value.ToString Then
                                If blnHighPerformanceReport = False Then
                                    xlWorkSheet.Cells(intExcelSheetRow, 20) = xlWorkSheet.Range("T" & intExcelSheetRow).Value + dgvPreviewExcel.Rows(intRow).Cells(19).Value
                                Else
                                    xlWorkSheet.Cells(intExcelSheetRow, 20) = xlWorkSheet.Range("T" & intExcelSheetRow).Value
                                End If
                                'xlWorkSheet.Cells(intExcelSheetRow, 20) = dgvPreviewExcel.Rows(intRow).Cells(19).Value
                            End If

                            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(20).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(0).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(0).Value.ToString Then
                                xlWorkSheet.Cells(intExcelSheetRow, 21) = xlWorkSheet.Range("U" & intExcelSheetRow).Value + dgvPreviewExcel.Rows(intRow).Cells(20).Value
                                'xlWorkSheet.Cells(intExcelSheetRow, 21) = dgvPreviewExcel.Rows(intRow).Cells(20).Value
                            End If

                            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(21).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(0).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(0).Value.ToString Then
                                xlWorkSheet.Cells(intExcelSheetRow, 22) = xlWorkSheet.Range("V" & intExcelSheetRow).Value + dgvPreviewExcel.Rows(intRow).Cells(21).Value
                                'xlWorkSheet.Cells(intExcelSheetRow, 22) = dgvPreviewExcel.Rows(intRow).Cells(21).Value
                            End If

                            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(22).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(0).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(0).Value.ToString Then
                                xlWorkSheet.Cells(intExcelSheetRow, 23) = xlWorkSheet.Range("W" & intExcelSheetRow).Value + dgvPreviewExcel.Rows(intRow).Cells(22).Value
                                'xlWorkSheet.Cells(intExcelSheetRow, 23) = dgvPreviewExcel.Rows(intRow).Cells(22).Value
                            End If

                        End If
                        intBaseRow += 1

                    ElseIf dgvPreviewExcel.Rows(intRow).Cells(1).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(1).Value.ToString Then
                        intExcelSheetRow += 1

                        If dgvPreviewExcel.Rows(intRow).Cells(1).Value.ToString <> "" Then
                            For intColumn = 0 To dgvPreviewExcel.Columns.Count - 1
                                DoEvents()
                                If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(intColumn).Value) = False Then
                                    xlWorkSheet.Cells(intExcelSheetRow, intColumn + 1) = CharPrior(intColumn) & Trim(dgvPreviewExcel.Rows(intRow).Cells(intColumn).Value)
                                End If
                            Next

                            If chkClearOrderPoint.Checked = True Then
                                xlWorkSheet.Cells(intExcelSheetRow, 10) = ""
                            End If
                            intBaseRow += 1
                        End If

                    End If
                End If
            End If


            tpbExport.Value = intRow + 1
        Next

        tpbExport.Value = dgvPreviewExcel.RowCount

        Dim strPrintAreaColumn1 As String = IIf(Trim(txtPrintAreaColumn1.Text) = "", "A", Trim(txtPrintAreaColumn1.Text))
        Dim strPrintAreaColumn2 As String = IIf(Trim(txtPrintAreaColumn1.Text) = "", "Z", Trim(txtPrintAreaColumn2.Text))
        Dim intZoom As Integer = IIf(Trim(txtZoom.Text) = "", 100, Trim(txtZoom.Text))

        With xlWorkSheet

            .Range(strRangeStart & "1", strRangeEnd & intExcelSheetRow + 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            .Range(strRangeStart & "1", strRangeEnd & intExcelSheetRow + 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            .Range(strRangeStart & "1", strRangeEnd & intExcelSheetRow + 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            .Range(strRangeStart & "1", strRangeEnd & intExcelSheetRow + 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            .Range(strRangeStart & "1", strRangeEnd & intExcelSheetRow + 1).Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
            .Range(strRangeStart & "1", strRangeEnd & intExcelSheetRow + 1).Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous

            If intRow <> 0 Then

                .Sort.SortFields.Clear()
                .Sort.SortFields.Add(.Columns(2))
                .Sort.SetRange(.Range(txtRangeStart.Text & "1:" & txtRangeEnd.Text & intExcelSheetRow))
                .Sort.Apply()

            End If

            .Range("F2", "F" & intExcelSheetRow).Cells.Interior.Color = RGB(202, 226, 199)
            .Range("I2", "I" & intExcelSheetRow).Cells.Interior.Color = RGB(254, 226, 227)
            .Range("J2", "J" & intExcelSheetRow).Cells.Interior.Color = RGB(228, 248, 254)
            .Range("K2", "K" & intExcelSheetRow).Cells.Interior.Color = RGB(245, 236, 154)
            .Range("X2", "Y" & intExcelSheetRow).NumberFormat = "#,###,###.00"


            .Range(strRangeStart & "1", strRangeEnd & "1").Select()
            xlApp.ActiveWindow.SplitColumn = 3
            xlApp.ActiveWindow.SplitRow = 1
            xlApp.ActiveWindow.FreezePanes = True

            For intColumn = 0 To dgvPreviewExcel.Columns.Count - 1
                DoEvents()
                .Cells(1, intColumn + 1) = Trim(dgvPreviewExcel.Columns(intColumn).HeaderText)
            Next

            .Range(strRangeStart & "1", strRangeEnd & "1").EntireColumn.AutoFit()
            .Range(strRangeStart & "1", strRangeEnd & "1").EntireRow.WrapText = True
            .Range(strRangeStart & "1", strRangeEnd & "1").EntireRow.VerticalAlignment = XlVAlign.xlVAlignCenter
            .Range(strRangeStart & "1", strRangeEnd & "1").Font.Bold = True
            .Range(strRangeStart & "1", strRangeEnd & "1").Cells.Interior.Color = RGB(173, 208, 239)
            .Range("A1").EntireRow.HorizontalAlignment = XlHAlign.xlHAlignCenter

            .Range("A1").ColumnWidth = 3
            .Range("B1").ColumnWidth = 15
            If blnHighPerformanceReport = False Then
                .Range("C1").ColumnWidth = 30
            Else
                .Range("C1").ColumnWidth = 10
            End If

            .Range("D1").ColumnWidth = 4
            .Range("E1").ColumnWidth = 3
            If blnHighPerformanceReport = False Then
                .Range("F1").ColumnWidth = 10
            Else
                .Range("F1").ColumnWidth = 30
            End If
            .Range("G1", "J1").ColumnWidth = 4
            .Range("K1", "W1").ColumnWidth = 4
            .Range("X1", "Y1").ColumnWidth = 6
            .Range("Z1").ColumnWidth = 10

            .Range("AA1").ColumnWidth = 10
            .Range("AC1").ColumnWidth = 10
            .Range("AF1").ColumnWidth = 10
            .Range("AK1").ColumnWidth = 5


            With .PageSetup
                .BottomMargin = 22
                .CenterFooter = "&P"
                .CenterHorizontally = True
                .FooterMargin = 11
                .LeftMargin = 0
                .Orientation = Orientation(cboPaperOrientation.Text)

                'Select Case cboPaperSize.Text
                '    Case "Legal"
                '        .PaperSize = XlPaperSize.xlPaperLegal
                '    Case "Letter"
                '        .PaperSize = XlPaperSize.xlPaperLetter
                'End Select


                .PaperSize = PaperSize(cboPaperSize.Text)

                .PrintArea = strPrintAreaColumn1 & "1:" & strPrintAreaColumn2 & intExcelSheetRow
                .PrintTitleRows = "$1:$1"
                .PrintTitleColumns = "$" & strPrintAreaColumn1 & ":$" & strPrintAreaColumn2
                .RightFooter = "&D&T"
                .RightMargin = 0
                .TopMargin = 0
                .Zoom = intZoom
            End With
        End With

        If chkMultiReport.Checked = True Then
            Dim xlZReportWorkSheet As Worksheet
            xlZReportWorkSheet = xlWorkSheet
            xlWorkSheet = xlWorkBook.Worksheets(2)
            xlWorkSheet.Name = "Discontinued Items"
            xlWorkSheet.Activate()
            ExportDiscontinuedItem(xlZReportWorkSheet, xlWorkSheet)
        End If

        'If chkPOImport.Checked = True Then
        '    Dim xlZReportWorkSheet As Worksheet
        '    xlZReportWorkSheet = xlWorkSheet

        '    xlWorkBook.Worksheets.Add(, 1, 1)
        '    xlWorkSheet = xlWorkBook.Worksheets(3)


        '    xlWorkSheet = xlWorkBook.Worksheets(3)
        '    xlWorkSheet.Name = "PO Import"
        '    xlWorkSheet.Activate()
        '    ExportDiscontinuedItem(xlZReportWorkSheet, xlWorkSheet)
        'End If

        xlWorkSheet = xlWorkBook.Worksheets(1)
        xlWorkSheet.Activate()

        xlWorkBook.SaveAs(DirPathEpicor & "\Surangel\" & strFileName & "_" & strDateTime & ".xlsx")
        xlApp.Visible = True



        'xlApp.Quit()

        Me.Cursor = Cursors.Default


        'Catch ex As Exception
        '    MessageBox.Show("Location Error: ExportDataToExcel, Error Message: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End

        'End Try

    End Sub


    Private Sub ExportHardwareReport(ByVal DataGridView As DataGridView)
        Dim strFileName As String = Path.GetFileNameWithoutExtension(strExportFullPath)
        Dim strDateTime As String = FormatDateTime(Now, DateFormat.ShortDate) & FormatDateTime(Now, DateFormat.ShortTime)
        Dim strVendorLookupName As String = ""
        Dim intRemoveRow As Integer = 0

        Me.Cursor = Cursors.WaitCursor
        strDateTime = Replace(strDateTime, "/", "")
        strDateTime = Replace(strDateTime, ":", "")
        Me.Cursor = Cursors.WaitCursor

        If xlApp Is Nothing Then
            MessageBox.Show("There's no Microsoft Excel installed in this machine")
            Exit Sub
        End If

        xlWorkBook = CType(xlApp.Workbooks.Add(), Excel.Workbook)
        xlWorkSheet = xlWorkBook.Worksheets(1)

        tpbExport.Maximum = dgvPreviewExcel.Rows.Count
        tpbExport.Visible = True
        dgvPreviewExcel.AllowUserToAddRows = False

        Dim intColumn As Integer = 0
        Dim intRow As Integer = 0

        Dim intBaseRow As Integer = 0
        Dim intExcelSheetRow As Integer = 2

        For intRow = 0 To dgvPreviewExcel.Rows.Count - 1
            lblProcessing.Text = "Status: Processing Row #:" & intRow

            DoEvents()

            If intRow = 0 Then
                For intColumn = 0 To dgvPreviewExcel.Columns.Count - 1
                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(intColumn).Value) = False Then
                        xlWorkSheet.Cells(intExcelSheetRow, intColumn + 1) = CharPrior(intColumn) & Trim(dgvPreviewExcel.Rows(intRow).Cells(intColumn).Value)
                    End If
                Next
            Else
                If dgvPreviewExcel.Rows(intRow).Cells(1).Value.ToString <> dgvPreviewExcel.Rows(intRow - 1).Cells(1).Value.ToString Then
                    If dgvPreviewExcel.Rows(intRow).Cells(1).Value.ToString <> "" Then
                        intExcelSheetRow += 1
                        For intColumn = 0 To dgvPreviewExcel.Columns.Count - 1
                            If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(intColumn).Value) = False Then
                                xlWorkSheet.Cells(intExcelSheetRow, intColumn + 1) = CharPrior(intColumn) & Trim(dgvPreviewExcel.Rows(intRow).Cells(intColumn).Value)
                            End If
                        Next
                    End If
                    'intBaseRow += 1

                ElseIf dgvPreviewExcel.Rows(intRow).Cells(1).Value.ToString = dgvPreviewExcel.Rows(intRow - 1).Cells(1).Value.ToString Then

                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(4).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(4).Value.ToString <> "" Then
                        xlWorkSheet.Range("E" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("E" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(4).Value.ToString)
                    End If

                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(5).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(5).Value.ToString <> "" Then
                        xlWorkSheet.Range("F" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("F" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(5).Value.ToString)
                    End If

                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(6).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(6).Value.ToString <> "" Then
                        xlWorkSheet.Range("G" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("G" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(6).Value.ToString)
                    End If

                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(7).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(7).Value.ToString <> "" Then
                        xlWorkSheet.Range("H" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("H" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(7).Value.ToString)
                    End If

                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(8).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(8).Value.ToString <> "" Then
                        xlWorkSheet.Range("I" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("I" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(8).Value.ToString)
                    End If


                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(13).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(13).Value.ToString <> "" Then
                        If Trim(dgvPreviewExcel.Columns(13).Name.ToString) = "StdPack" Then
                            xlWorkSheet.Range("N" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("N" & intExcelSheetRow).Value) '+ Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(13).Value.ToString)
                        Else
                            xlWorkSheet.Range("N" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("N" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(13).Value.ToString)
                        End If
                    End If

                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(14).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(14).Value.ToString <> "" Then
                        xlWorkSheet.Range("O" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("O" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(14).Value.ToString)
                    End If


                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(15).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(15).Value.ToString <> "" Then
                        xlWorkSheet.Range("P" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("P" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(15).Value.ToString)
                    End If


                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(22).Value) = False And IsNumeric(dgvPreviewExcel.Rows(intRow).Cells(22).Value) = True And dgvPreviewExcel.Rows(intRow).Cells(22).Value.ToString <> "" Then
                        xlWorkSheet.Range("W" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("W" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(22).Value.ToString)
                    End If

                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(23).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(23).Value.ToString <> "" Then
                        xlWorkSheet.Range("X" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("X" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(23).Value.ToString)
                    End If

                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(24).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(24).Value.ToString <> "" Then
                        xlWorkSheet.Range("Y" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("Y" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(24).Value.ToString)
                    End If

                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(25).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(25).Value.ToString <> "" Then
                        xlWorkSheet.Range("Z" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("Z" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(25).Value.ToString)
                    End If

                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(26).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(26).Value.ToString <> "" Then
                        xlWorkSheet.Range("AA" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("AA" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(26).Value.ToString)
                    End If

                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(27).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(27).Value.ToString <> "" Then
                        xlWorkSheet.Range("AB" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("AB" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(27).Value.ToString)
                    End If

                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(28).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(28).Value.ToString <> "" Then
                        xlWorkSheet.Range("AC" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("AC" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(28).Value.ToString)
                    End If

                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(29).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(29).Value.ToString <> "" Then
                        xlWorkSheet.Range("AD" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("AD" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(29).Value.ToString)
                    End If

                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(30).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(30).Value.ToString <> "" Then
                        xlWorkSheet.Range("AE" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("AE" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(30).Value.ToString)
                    End If

                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(31).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(31).Value.ToString <> "" Then
                        xlWorkSheet.Range("AF" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("AF" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(31).Value.ToString)
                    End If

                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(32).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(32).Value.ToString <> "" Then
                        xlWorkSheet.Range("AG" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("AG" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(32).Value.ToString)
                    End If
                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(33).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(33).Value.ToString <> "" Then
                        xlWorkSheet.Range("AH" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("AH" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(33).Value.ToString)
                    End If
                    If IsDBNull(dgvPreviewExcel.Rows(intRow).Cells(34).Value) = False And dgvPreviewExcel.Rows(intRow).Cells(34).Value.ToString <> "" Then
                        xlWorkSheet.Range("AI" & intExcelSheetRow).Value = Convert.ToDouble(xlWorkSheet.Range("AI" & intExcelSheetRow).Value) + Convert.ToDouble(dgvPreviewExcel.Rows(intRow).Cells(34).Value.ToString)
                    End If

                    'intBaseRow += 1
                End If

            End If
            tpbExport.Value = intRow + 1
        Next


        Dim strPrintAreaColumn1 As String = IIf(Trim(txtPrintAreaColumn1.Text) = "", "A", Trim(txtPrintAreaColumn1.Text))
        Dim strPrintAreaColumn2 As String = IIf(Trim(txtPrintAreaColumn1.Text) = "", "Z", Trim(txtPrintAreaColumn2.Text))
        Dim intZoom As Integer = IIf(Trim(txtZoom.Text) = "", 100, Trim(txtZoom.Text))

        With xlWorkSheet
            .Range(strRangeStart & "1", strRangeEnd & "1").Select()
            xlApp.ActiveWindow.SplitRow = 1
            xlApp.ActiveWindow.FreezePanes = True

            For intColumn = 0 To dgvPreviewExcel.Columns.Count - 1
                .Cells(1, intColumn + 1) = Trim(dgvPreviewExcel.Columns(intColumn).HeaderText)
            Next

            lblProcessing.Text = "Status: Please wait while sorting by SKUs..."
            .Range(strRangeStart & "1", strRangeEnd & intRow).Sort(Key1:= .Range(Trim(txtSortLevel1.Text) & "1"),
                        Order1:=Excel.XlSortOrder.xlAscending, Header:=Excel.XlYesNoGuess.xlYes,
                        MatchCase:=True, Orientation:=Excel.XlSortOrientation.xlSortColumns, DataOption1:=XlSortDataOption.xlSortTextAsNumbers)

            'lblProcessing.Text = "Status: Creating the sub totals..."
            'Dim ColumnFields() As Integer = {5, 6, 7, 8, 14, 15, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35}

            '.Range(strRangeStart & "1", strRangeEnd & intRow + 1).Subtotal(2, XlConsolidationFunction.xlSum, ColumnFields, True, False, XlSummaryRow.xlSummaryBelow)

            .Range(strRangeStart & "1", strRangeEnd & "1").EntireColumn.AutoFit()
            .Range(strRangeStart & "1", strRangeEnd & "1").EntireRow.WrapText = True
            .Range(strRangeStart & "1", strRangeEnd & "1").EntireRow.VerticalAlignment = XlVAlign.xlVAlignCenter
            .Range(strRangeStart & "1", strRangeEnd & "1").Font.Bold = True
            .Range(strRangeStart & "1", strRangeEnd & "1").Font.Color = RGB(255, 255, 255)
            .Range(strRangeStart & "1", strRangeEnd & "1").Cells.Interior.Color = RGB(33, 138, 184)

            .Name = "Combined"

            lblProcessing.Text = "Status: Please wait while setting up the page printing..."
            With .PageSetup
                .BottomMargin = 22
                .CenterFooter = "&P"
                .CenterHorizontally = True
                .FooterMargin = 11
                .LeftMargin = 0
                .Orientation = Orientation(cboPaperOrientation.Text)
                .PaperSize = PaperSize(cboPaperSize.Text)
                .PrintArea = strPrintAreaColumn1 & "1:" & strPrintAreaColumn2 & intExcelSheetRow
                .PrintTitleRows = "$1:$1"
                .PrintTitleColumns = "$" & strPrintAreaColumn1 & ":$" & strPrintAreaColumn2
                .RightFooter = "&D&T"
                .RightMargin = 0
                .TopMargin = 0
                .Zoom = intZoom

            End With
        End With

        'Try

        '    lblProcessing.Text = "Status: Creating the sub total page ..."
        '    'SubTotal Worksheet
        '    Dim intXlSubTotalRow As Integer = 0
        '    Dim isEmpty As Boolean = False
        '    Dim strCellValue As String = ""
        '    Dim isTotalFound As Boolean = False
        '    Dim xlWSSubTotal As Excel.Worksheet
        '    xlWSSubTotal = xlWorkBook.Worksheets(2)
        '    With xlWSSubTotal
        '        For intColumn = 1 To dgvPreviewExcel.ColumnCount
        '            .Cells(1, intColumn) = xlWorkSheet.Cells(1, intColumn)
        '        Next

        '        lblProcessing.Text = "Status: Processing the totals..."
        '        intRow = 2
        '        intXlSubTotalRow = 2

        '        Do While isEmpty = False
        '            Try
        '                strCellValue = LCase(Trim(xlWorkSheet.Range("B" & intRow).Value))
        '                isEmpty = IIf(strCellValue = "", True, False)

        '                lblProcessing.Text = "Processing the totals..." & intRow - 1 & "/" & dgvPreviewExcel.RowCount & " CellValue = " & strCellValue

        '                If isEmpty = False Then
        '                    If Trim(Mid(strCellValue, IIf(strCellValue.Length - 5 <= 0, strCellValue.Length, strCellValue.Length - 5), strCellValue.Length)) = "total" Then
        '                        For intColumn = 1 To dgvPreviewExcel.ColumnCount
        '                            .Cells(intXlSubTotalRow, intColumn) = xlWorkSheet.Cells(intRow, intColumn)
        '                        Next
        '                        If Trim(Mid(strCellValue, 1, 5)) <> "grand" Then
        '                            .Cells(intXlSubTotalRow, 3) = xlWorkSheet.Cells(intRow - 1, 3)
        '                            xlWorkSheet.Cells(intRow, 3) = xlWorkSheet.Cells(intRow - 1, 3)
        '                        End If

        '                        intXlSubTotalRow += 1
        '                    End If
        '                End If

        '                'If intRow >= dgvPreviewExcel.RowCount Then
        '                '    Exit Do
        '                'End If

        '            Catch ex As Exception
        '                MessageBox.Show(ex.Message)
        '                GoTo SaveFile
        '            End Try

        '            intRow += 1
        '        Loop


        '.Range(strRangeStart & "1", strRangeEnd & "1").EntireColumn.AutoFit()
        '.Range(strRangeStart & "1", strRangeEnd & "1").EntireRow.WrapText = True
        '.Range(strRangeStart & "1", strRangeEnd & "1").EntireRow.VerticalAlignment = XlVAlign.xlVAlignCenter
        '.Range(strRangeStart & "1", strRangeEnd & "1").Font.Bold = True
        '.Range(strRangeStart & "1", strRangeEnd & "1").Font.Color = RGB(255, 255, 255)
        '.Range(strRangeStart & "1", strRangeEnd & "1").Cells.Interior.Color = RGB(33, 138, 184)
        '.Name = "SubTotal"

        'End With


        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        '    GoTo SaveFile

        'End Try

        If chkMultiReport.Checked = True Then
            Dim xlZReportWorkSheet As Worksheet
            xlZReportWorkSheet = xlWorkSheet
            xlWorkSheet = xlWorkBook.Worksheets(2)
            xlWorkSheet.Name = "Discontinued Items"
            xlWorkSheet.Activate()
            ExportDiscontinuedItem(xlZReportWorkSheet, xlWorkSheet)
        End If

SaveFile:

        xlWorkSheet = xlWorkBook.Worksheets(1)
        xlWorkSheet.Activate()

        lblProcessing.Text = "Status: Please wait while saving the file..."

        xlWorkBook.SaveAs(DirPathEpicor & "\Surangel\" & IIf(strVendorLookupName <> "", strVendorLookupName, strFileName) & "_" & strDateTime & ".xlsx")
        xlApp.Visible = True

        dgvPreviewExcel.Cursor = Cursors.Default

        lblProcessing.Text = "Status: Finished!"
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub ExportExpiredReportToExcel(ByVal DataGridView As DataGridView)
        Me.Cursor = Cursors.WaitCursor
        Dim strFileName As String = Path.GetFileNameWithoutExtension(strExportFullPath)
        Dim strDateTime As String = FormatDateTime(Now, DateFormat.ShortDate) & FormatDateTime(Now, DateFormat.ShortTime)

        If xlApp Is Nothing Then
            MessageBox.Show("There's no Microsoft Excel installed in this machine")
            Exit Sub
        End If

        xlWorkBook = CType(xlApp.Workbooks.Add(), Office.Interop.Excel.Workbook)
        xlWorkSheet = xlWorkBook.Worksheets(1)

        tpbExport.Maximum = dgvPreviewExcel.Rows.Count
        tpbExport.Visible = True
        dgvPreviewExcel.AllowUserToAddRows = False

        Dim intColumn As Integer = 0
        Dim intRow As Integer = 0
        Dim intBaseRow As Integer = 0
        Dim intExcelSheetRow As Integer = 2

        Dim lngExpirationDate As Long = 0

        'Header
        With xlWorkSheet
            .Cells(intRow + 1, 1) = "=Today()"
            For intColumn = 0 To dgvPreviewExcel.Columns.Count - 1
                .Cells(1, intColumn + 2) = Trim(dgvPreviewExcel.Columns(intColumn).HeaderText)
            Next

            For intRow = 0 To dgvPreviewExcel.Rows.Count - 1
                DoEvents()
                .Cells(intRow + 2, 1) = intRow + 1
                .Cells(intRow + 2, 2) = Trim(dgvPreviewExcel.Rows(intRow).Cells("SKU").Value)
                .Cells(intRow + 2, 3) = Trim(dgvPreviewExcel.Rows(intRow).Cells("ItemDescription").Value)
                .Cells(intRow + 2, 4) = Trim(dgvPreviewExcel.Rows(intRow).Cells("Pack").Value)
                .Cells(intRow + 2, 5) = Trim(dgvPreviewExcel.Rows(intRow).Cells("PURUOM").Value)
                .Cells(intRow + 2, 6) = Trim(dgvPreviewExcel.Rows(intRow).Cells("UPC").Value)
                .Cells(intRow + 2, 7) = Trim(dgvPreviewExcel.Rows(intRow).Cells("ExpirationDate").Value)

                'lngExpirationDate = Convert.ToString(DateDiff(DateInterval.Day, dgvPreviewExcel.Rows(intRow).Cells("ExpirationDate").Value, Date.Today))
                'If lngExpirationDate <= 30 Then
                '.Range("G" & intExcelSheetRow).FormatConditions.Add(XlFormatConditionType.xlCellValue,

                '.Cells.Interior.Color = RGB(225, 153, 153)
                'End If

                .Cells(intRow + 2, 8) = Trim(dgvPreviewExcel.Rows(intRow).Cells("QOHST1").Value)
                .Cells(intRow + 2, 9) = Trim(dgvPreviewExcel.Rows(intRow).Cells("QOHST3").Value)
                .Cells(intRow + 2, 10) = Trim(dgvPreviewExcel.Rows(intRow).Cells("QOHST7").Value)
                .Cells(intRow + 2, 11) = Trim(dgvPreviewExcel.Rows(intRow).Cells("QOHTotal").Value)
                .Cells(intRow + 2, 12) = Trim(dgvPreviewExcel.Rows(intRow).Cells("OrderPoint").Value)
                .Cells(intRow + 2, 13) = Trim(dgvPreviewExcel.Rows(intRow).Cells("QOO").Value)
                .Cells(intRow + 2, 14) = Trim(dgvPreviewExcel.Rows(intRow).Cells("SalesUnits1").Value)
                .Cells(intRow + 2, 15) = Trim(dgvPreviewExcel.Rows(intRow).Cells("SalesUnits2").Value)
                .Cells(intRow + 2, 16) = Trim(dgvPreviewExcel.Rows(intRow).Cells("SalesUnits3").Value)
                .Cells(intRow + 2, 17) = Trim(dgvPreviewExcel.Rows(intRow).Cells("SalesUnits4").Value)
                .Cells(intRow + 2, 18) = Trim(dgvPreviewExcel.Rows(intRow).Cells("RetailPrice").Value)
                .Cells(intRow + 2, 19) = Trim(dgvPreviewExcel.Rows(intRow).Cells("AvgCost").Value)
                .Cells(intRow + 2, 20) = Trim(dgvPreviewExcel.Rows(intRow).Cells("GMROI").Value)
                .Cells(intRow + 2, 21) = Trim(dgvPreviewExcel.Rows(intRow).Cells("ReplCost").Value)
                .Cells(intRow + 2, 22) = Trim(dgvPreviewExcel.Rows(intRow).Cells("DateOfLastSale").Value)
                .Cells(intRow + 2, 23) = Trim(dgvPreviewExcel.Rows(intRow).Cells("DepartmentCode").Value)
                .Cells(intRow + 2, 24) = Trim(dgvPreviewExcel.Rows(intRow).Cells("MFGPart#").Value)
                .Cells(intRow + 2, 25) = Trim(dgvPreviewExcel.Rows(intRow).Cells("SalesUnitYTD").Value)
                .Cells(intRow + 2, 26) = Trim(dgvPreviewExcel.Rows(intRow).Cells("LastYearUnits").Value)
                .Cells(intRow + 2, 27) = Trim(dgvPreviewExcel.Rows(intRow).Cells("RetailPriceChange").Value)
                .Cells(intRow + 2, 28) = Trim(dgvPreviewExcel.Rows(intRow).Cells("VendorName").Value)
                lblProcessing.Text = "Status: Processing Row #:" & intRow + 1
                intExcelSheetRow += 1
            Next


            '.Range("G1", "G" & intExcelSheetRow)


            Dim strPrintAreaColumn1 As String = IIf(Trim(txtPrintAreaColumn1.Text) = "", "A", Trim(txtPrintAreaColumn1.Text))
            Dim strPrintAreaColumn2 As String = IIf(Trim(txtPrintAreaColumn1.Text) = "", "Z", Trim(txtPrintAreaColumn2.Text))
            Dim intZoom As Integer = IIf(Trim(txtZoom.Text) = "", 100, Trim(txtZoom.Text))

            .Range(strRangeStart & "1", strRangeEnd & "1").Select()
            xlApp.ActiveWindow.SplitRow = 1
            xlApp.ActiveWindow.FreezePanes = True

            .Range(strRangeStart & "1", strRangeEnd & "1").EntireColumn.AutoFit()
            .Range(strRangeStart & "1", strRangeEnd & "1").EntireRow.WrapText = True
            .Range(strRangeStart & "1", strRangeEnd & "1").EntireRow.VerticalAlignment = XlVAlign.xlVAlignCenter
            .Range(strRangeStart & "1", strRangeEnd & "1").Font.Bold = True

            '.Range(strRangeStart & "1", strRangeEnd & "1").Cells.Interior.Color = RGB(173, 208, 239)
            '.Range("A1").EntireRow.HorizontalAlignment = XlHAlign.xlHAlignCenter

            '.Range("A1").ColumnWidth = 3
            '.Range("B1").ColumnWidth = 15
            '.Range("C1").ColumnWidth = 30
            '.Range("D1").ColumnWidth = 4
            '.Range("E1").ColumnWidth = 3
            '.Range("F1").ColumnWidth = 10
            '.Range("G1", "J1").ColumnWidth = 4
            '.Range("K1", "W1").ColumnWidth = 4
            '.Range("X1", "Y1").ColumnWidth = 6
            '.Range("Z1").ColumnWidth = 10

            '.Range("AA1").ColumnWidth = 10
            '.Range("AC1").ColumnWidth = 10
            '.Range("AF1").ColumnWidth = 10
            '.Range("AK1").ColumnWidth = 5

            With .PageSetup
                .BottomMargin = 22
                .CenterFooter = "&P"
                .CenterHorizontally = True
                .FooterMargin = 11
                .LeftMargin = 0
                .Orientation = Orientation(cboPaperOrientation.Text)

                'Select Case cboPaperSize.Text
                '    Case "Legal"
                '        .PaperSize = XlPaperSize.xlPaperLegal
                '    Case "Letter"
                '        .PaperSize = XlPaperSize.xlPaperLetter
                'End Select

                .PaperSize = PaperSize(cboPaperSize.Text)

                .PrintArea = strPrintAreaColumn1 & "1:" & strPrintAreaColumn2 & intExcelSheetRow
                .PrintTitleRows = "$1:$1"
                .PrintTitleColumns = "$" & strPrintAreaColumn1 & ":$" & strPrintAreaColumn2
                .RightFooter = "&D&T"
                .RightMargin = 0
                .TopMargin = 0
                .Zoom = intZoom
            End With
        End With


        xlWorkSheet = xlWorkBook.Worksheets(1)
        xlWorkSheet.Activate()

        strFileName = "MonthlyExpiredProductReport"
        strDateTime = Replace(strDateTime, "/", "")
        strDateTime = Replace(strDateTime, ":", "")
        xlWorkBook.SaveAs(DirPathEpicor & "\Surangel\" & strFileName & "_" & strDateTime & ".xlsx")
        xlApp.Visible = True

        Me.Cursor = Cursors.Default

    End Sub

    Private Function GetColumnNumber(ByVal DataGridView As DataGridView, ByVal FindString As String) As Integer
        Dim intColumn As Integer = 0

        For intColumn = 0 To dgvPreviewExcel.Columns.Count - 1
            DoEvents()
            If dgvPreviewExcel.Columns(intColumn).Name.Contains(FindString) = True Then
                Exit For
            End If
        Next
        Return intColumn
    End Function

    Private Function RemoveNumeric(ByVal Value As String) As String
        Dim intLoop As Integer
        Dim strNewValue As String = ""

        For intLoop = 1 To Value.Length
            If IsNumeric(Mid(Value, intLoop, 1)) = False Then
                strNewValue &= Mid(Value, intLoop, 1)
            End If
        Next
        Return strNewValue

    End Function

    Private Sub ProcessMonthHeaders()
        Dim strPeriodHeader(11) As String
        Dim strPeriod(11) As String
        Dim intLoop As Integer = 0

        'Exit Sub

        Dim intCurrentMonth As Integer = Month(FormatDateTime(Now, DateFormat.ShortDate))
        Dim intColumnStart = GetColumnNumber(dgvPreviewExcel, "Sales Units")

        Me.Cursor = Cursors.WaitCursor
        For intLoop = 0 To 11
            DoEvents()
            strPeriod(intLoop) = MonthName(intCurrentMonth, True)
            With dgvPreviewExcel
                If .RowCount <> 0 Then
                    .Columns(intColumnStart).HeaderText = Replace(.Columns(intColumnStart).HeaderText, "Sales Units", "")
                    .Columns(intColumnStart).HeaderText = Replace(.Columns(intColumnStart).HeaderText, "Period", "")
                    .Columns(intColumnStart).HeaderText = RemoveNumeric(.Columns(intColumnStart).HeaderText)
                    .Columns(intColumnStart).HeaderText &= " " & strPeriod(intLoop)
                End If
            End With

            If intCurrentMonth >= 2 Then
                intCurrentMonth -= 1
            Else
                intCurrentMonth += 11
            End If
            intColumnStart += 1
        Next

        Me.Cursor = Cursors.Default

        'MessageBox.Show("ProcessMonthHeaders")

    End Sub

    Private Function GetSheetName(ByVal FileName As String)
        Dim strSheetName As String = ""

        Return strSheetName
    End Function

    Private Sub ConvertCVSToExcel(ByVal SourceFile As String)
        Dim dt As New System.Data.DataTable
        Dim csvReader As New FileIO.TextFieldParser(SourceFile)

        csvReader.TextFieldType = FileIO.FieldType.Delimited
        csvReader.SetDelimiters(",")

        Dim rowData As String
        Dim currentField As String
        Dim currentRow As String()

        Dim intColumn As Integer = 0
        Dim intRow As Integer = 0

        lblProcessing.Text = "Status: Please white while processing..."
        DoEvents()
        With dt
            For Each currentField In csvReader.ReadFields
                .Columns.Add(currentField)
            Next

            While Not csvReader.EndOfData
                currentRow = csvReader.ReadFields()
                .Rows.Add()
                For Each rowData In currentRow
                    .Rows(intRow).Item(intColumn) = rowData
                    intColumn += 1
                Next

                intColumn = 0
                intRow += 1
            End While

            .DefaultView.Sort = "SKU, St"

            '.Select.OrderBy("SKU")
            '.Select("", "SKU, St")

        End With

        dgvPreviewExcel.DataSource = dt


        'With dgvPreviewExcel

        '    For Each currentField In csvReader.ReadFields
        '        .Columns.Add(currentField, currentField)
        '    Next

        '    While Not csvReader.EndOfData
        '        currentRow = csvReader.ReadFields()
        '        .Rows.Add()
        '        For Each rowData In currentRow
        '            .Rows(intRow).Cells(intColumn).Value = rowData
        '            intColumn += 1
        '        Next
        '        intColumn = 0
        '        intRow += 1
        '    End While
        'End With

    End Sub

    Private Function ExtractItemNumber(ByVal Number As String, ByVal Starting As Integer, ByVal Ending As Integer, ByVal LoopStep As Integer)

        Me.Cursor = Cursors.WaitCursor
        Dim strItemNumber As String = ""
        Dim intLenght As Integer = 0
        Dim intLoop As Integer = 0

        Number = Trim(Number)
        For intLoop = Starting To Ending Step LoopStep
            If IsNumeric(Mid(Number, intLoop, 1)) = True Then
                If LoopStep > 0 Then
                    strItemNumber &= Mid(Number, intLoop, 1)
                Else
                    strItemNumber = Mid(Number, intLoop, 1) & strItemNumber
                End If
            ElseIf Mid(Number, intLoop, 1) = "-" Then
                Exit For
            End If
        Next

        Me.Cursor = Cursors.Default

        Return strItemNumber

    End Function

    Private Function RemovePriorAlphaItemNumber(ByVal Number As String)

        Me.Cursor = Cursors.WaitCursor
        Dim strItemNumber As String = ""
        Dim intLenght As Integer = 0
        Dim intLoop As Integer = 0

        'strItemNumber = Trim(Number)
        For intLoop = 1 To Number.Length
            If IsNumeric(Mid(Number, intLoop, 1)) = True Or Mid(Number, intLoop, 1) = "-" Then
                strItemNumber &= Mid(Number, intLoop, 1)
            End If
        Next
        'MessageBox.Show("1. strItemNumber = " & strItemNumber & " Number = " & Number)

        If Mid(strItemNumber, 1, 1) = "-" Then
            strItemNumber = Mid(strItemNumber, 2, strItemNumber.Length)
        End If

        'MessageBox.Show("2. strItemNumber = " & strItemNumber & " Number = " & Number)

        If strItemNumber.Length > 0 Then
            If Mid(strItemNumber, strItemNumber.Length, 1) = "-" Then
                strItemNumber = Mid(strItemNumber, 1, strItemNumber.Length - 1)
            End If
        End If
        Me.Cursor = Cursors.Default

        'MessageBox.Show("3. strItemNumber = " & strItemNumber & " Number = " & Number)
        Return strItemNumber

    End Function

    Private Sub GridViewColumnExpiredDateReport()

        With dgvPreviewExcel
            .Columns.Add("SKU", "SKU")
            .Columns.Add("ItemDescription", "Item Description")
            .Columns.Add("Pack", "Pack")
            .Columns.Add("PURUOM", "PUR UOM")
            .Columns.Add("UPC", "UPC")
            .Columns.Add("ExpirationDate", "Expiration Date")
            .Columns.Add("QOHST1", "QOH ST1")
            .Columns.Add("QOHST3", "QOH ST3")
            .Columns.Add("QOHST7", "QOH ST7")
            .Columns.Add("QOHTotal", "QOHTotal")
            .Columns.Add("OrderPoint", "Order Point")
            .Columns.Add("QOO", "QOO")
            '.Columns.Add("Total", "Total")
            .Columns.Add("SalesUnits1", "SalesUnits1")
            .Columns.Add("SalesUnits2", "SalesUnits2")
            .Columns.Add("SalesUnits3", "SalesUnits3")
            .Columns.Add("SalesUnits4", "SalesUnits4")
            .Columns.Add("RetailPrice", "Retail Price")
            .Columns.Add("AvgCost", "Average Cost")
            .Columns.Add("GMROI", "GMROI")
            .Columns.Add("ReplCost", "Repl Cost")
            .Columns.Add("DateOfLastSale", "Date Of Last Sale")
            .Columns.Add("DateOfLastReceipt", "Date Of Last Receipt")
            .Columns.Add("DepartmentCode", "Department Code")
            .Columns.Add("MFGPart#", "MFG Part #")
            .Columns.Add("SalesUnitYTD", "Sales Unit YTD")
            .Columns.Add("LastYearUnits", "Last Year Units")
            .Columns.Add("RetailPriceChange", "Retail PriceChange")
            .Columns.Add("VendorName", "Vendor Name")
            '.Columns.Add("UPCPrimary", "UPC Primary")
        End With
    End Sub

    Private Function PaperSize(ByVal strPaperSize As String) As XlPaperSize
        Dim PrintPaperSize As XlPaperSize

        Select Case strPaperSize
            Case "Legal"
                PrintPaperSize = XlPaperSize.xlPaperLegal
            Case "Letter"
                PrintPaperSize = XlPaperSize.xlPaperLetter
        End Select

        Return PrintPaperSize
    End Function

    Private Function Orientation(ByVal PaperOrientation As String) As XlPageOrientation
        Dim objOrientation As XlPageOrientation

        Select Case PaperOrientation
            Case "Landscape" : objOrientation = XlPageOrientation.xlLandscape
            Case "Portrait" : objOrientation = XlPageOrientation.xlPortrait
        End Select
        Return objOrientation

    End Function

    Private Sub Preview(ByVal FileName As String, Optional ByVal FilePath As String = "")

        'Try

        Me.Cursor = Cursors.WaitCursor
        Dim strTableName As String = ""
        Dim isCSVFormat As Boolean = False
        Dim strFileName As String = Path.GetFileName(FileName)

        Dim xlApp As New Office.Interop.Excel.Application

        Dim xlConnection As New System.Data.OleDb.OleDbConnection
        Dim xlDataset As New System.Data.DataSet
        Dim xlDataAdapter As New System.Data.OleDb.OleDbDataAdapter
        Dim xlCommand As New System.Data.OleDb.OleDbCommand

        'Dim strUPC As String = ""
        'Dim strQty As String = ""
        'Dim strExpirationDate As String = ""

        strExportFullPath = FileName

        DoEvents()

        lblProcessing.Text = "Status: Please wait while loading the data..."
        If blnCombineReport = False Then
            dgvPreviewExcel.Rows.Clear()
            dgvPreviewExcel.Columns.Clear()
        End If

        If Mid(FileName, FileName.Length - 2, 3) = "csv" Then
            ConvertCVSToExcel(FileName)

        ElseIf Mid(FileName, FileName.Length - 2, 3) = "txt" Then
            'Dim TxtReader As TextReader = New StreamReader(FilePath & "\" & FileName)
            'Dim recordLine As String = ""

            ''Add GridView Column
            ''GridViewColumnExpiredDateReport()

            'Do Until recordLine Is Nothing
            '    If TxtReader.Peek = -1 Then
            '        Exit Do
            '    End If

            '    recordLine = TxtReader.ReadLine()
            '    Dim list As IList(Of String) = New List(Of String)(recordLine.Split(New String() {","}, StringSplitOptions.None))

            '    strUPC = Trim(list(0))
            '    strQty = Trim(list(1))
            '    strExpirationDate = Trim(list(2))


            '    'MessageBox.Show("strUPC: " & strUPC & vbCrLf & "strQty: " & strQty & vbCrLf & "strExpirationDate: " & strExpirationDate)
            'Loop

            GoTo ExitTextProcess

        Else
            'xlConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source= " & FilePath & "\" & FileName & ";Extended Properties = 'Excel 12.0 Xml;HDR=YES;'")
            xlConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source= " & FilePath & ";Extended Properties = 'Excel 12.0 Xml;HDR=YES;'")
            xlConnection.Open()
            xlCommand.Connection = xlConnection
            xlCommand.CommandType = CommandType.Text

            Try
                xlCommand.CommandText = "select * from [Sheet1$] order by [Item Number], Store"
                xlCommand.ExecuteNonQuery()

            Catch ex As Exception
                If blnProductExpirationReport = True Then
                    xlCommand.CommandText = "select * from [Sheet1$] order by SKU"
                Else
                    xlCommand.CommandText = "select * from [Sheet1$] order by SKU, St"
                End If
                xlCommand.ExecuteNonQuery()
            End Try

            xlDataAdapter.SelectCommand = xlCommand
            xlDataAdapter.TableMappings.Add("Table", "Sheet1")
            xlDataAdapter.Fill(xlDataset)

            xlConnection.Close()


            DoEvents()
            With dgvPreviewExcel
                .DataSource = Nothing
                .DataSource = xlDataset
                .DataMember = "Sheet1"

                'Fixed columns for High Performance Report
                If .Columns.Count = 26 Then
                    blnHighPerformanceReport = True
                End If


                If blnProductExpirationReport = True Then
                    lblColumns.Text = "Columns: " & dgvPreviewExcel.Columns.Count
                    lblRows.Text = "Rows: " & dgvPreviewExcel.Rows.Count

                    xlConnection.Close()
                    lblProcessing.Text = "Status: Ready to process"
                    butProcess.Enabled = True
                    If blnCombineReport = True Then
                        lsvFiles.Enabled = True
                    End If
                    Me.Cursor = Cursors.Default

                    Exit Sub
                End If

                .Columns.Add("OrderPointCopy", "Order Point Copy")
                .Columns.Add("ItemNumber0", "ItemNumber0")
                .Columns.Add("ItemNumber1", "ItemNumber1")
                .Columns.Add("ItemNumber2", "ItemNumber2")

                .Columns.Add("tmpLastSale", "tmpLastSale")
                .Columns.Add("tmpLastReceipt", "tmpLastReceipt")


                'MessageBox.Show(".Rows.Count  =" & .Rows.Count)
                'Exit Sub

                For intloop = 0 To .RowCount - 1
                    .Rows(intloop).Cells("OrderPointCopy").Value = .Rows(intloop).Cells(9).Value

                    If IsDBNull(.Rows(intloop).Cells(1).Value) = False And .Rows(intloop).Cells(1).Value.ToString <> "" Then
                        .Rows(intloop).Cells("ItemNumber0").Value = RemovePriorAlphaItemNumber(.Rows(intloop).Cells(1).Value)
                    End If

                    If Convert.ToString(.Rows(intloop).Cells(25).Value) <> "" Then
                        If IsDate(.Rows(intloop).Cells(25).Value) = True Then
                            .Rows(intloop).Cells("tmpLastSale").Value = FormatDateTime(.Rows(intloop).Cells(25).Value.ToString, DateFormat.ShortDate)
                        End If
                    End If

                    If .Rows(intloop).Cells(26).Value.ToString <> "" Then
                        If IsDate(.Rows(intloop).Cells(26).Value) = True Then
                            .Rows(intloop).Cells("tmpLastReceipt").Value = FormatDateTime(.Rows(intloop).Cells(26).Value.ToString, DateFormat.ShortDate)
                        End If
                    End If

                Next

                For intloop = 0 To .RowCount - 1
                    If .Rows(intloop).Cells("ItemNumber0").Value <> "" Then
                        .Rows(intloop).Cells("ItemNumber1").Value = ExtractItemNumber(.Rows(intloop).Cells("ItemNumber0").Value.ToString, 1, .Rows(intloop).Cells("ItemNumber0").Value.ToString.Length, 1)
                    End If

                    If .Rows(intloop).Cells("tmpLastSale").Value <> "" Then
                        .Rows(intloop).Cells(26).Value = .Rows(intloop).Cells("tmpLastSale").Value
                    End If
                    If .Rows(intloop).Cells("tmpLastReceipt").Value <> "" Then
                        .Rows(intloop).Cells(25).Value = .Rows(intloop).Cells("tmpLastReceipt").Value
                    End If

                Next

                For intloop = 0 To .RowCount - 1
                    If .Rows(intloop).Cells("ItemNumber0").Value <> "" Then
                        .Rows(intloop).Cells("ItemNumber2").Value = ExtractItemNumber(.Rows(intloop).Cells("ItemNumber0").Value, .Rows(intloop).Cells("ItemNumber0").Value.ToString.Length, 1, -1)
                    End If
                Next

            End With

            Dim intCol As Integer = 0
            Dim intRow As Integer = 0
            Select Case intDGVPreview
                Case 0
                    Me.Controls.Add(dgvTemp0)
                    With dgvTemp0
                        .DataSource = Nothing
                        .DataSource = xlDataset
                        .DataMember = "Sheet1"
                        .Visible = False
                        .Left = 0
                        .Top = 0
                        .Width = 500
                        .Height = 500
                        .Columns.RemoveAt(0)
                        .AllowUserToAddRows = False

                    End With
                    intCombinedReport += 1

                Case 1
                    Me.Controls.Add(dgvTemp1)
                    With dgvTemp1
                        .DataSource = Nothing
                        .DataSource = xlDataset
                        .DataMember = "[Sheet1$]"
                        .Visible = False
                        .Left = 510
                        .Top = 200
                        .Width = 500
                        .Height = 500
                        .Columns.RemoveAt(0)
                        .AllowUserToAddRows = False
                    End With
                    intCombinedReport += 1

                Case 2
                    With dgvTemp2
                        .DataSource = Nothing
                        .DataSource = xlDataset
                        .DataMember = "[Sheet1$]"
                        .Visible = False
                        .AllowUserToAddRows = False
                    End With
                    Me.Controls.Add(dgvTemp2)
                    intCombinedReport += 1

                Case 3
                    Me.Controls.Add(dgvTemp3)
                    With dgvTemp3
                        .DataSource = Nothing
                        .DataSource = xlDataset
                        .DataMember = "[Sheet1$]"
                        .Visible = False
                        .AllowUserToAddRows = False
                    End With
                    intCombinedReport += 1

                Case 4
                    Me.Controls.Add(dgvTemp4)
                    With dgvTemp4
                        .DataSource = Nothing
                        .DataSource = xlDataset
                        .DataMember = "[Sheet1$]"
                        .AllowUserToAddRows = False
                    End With
                    intCombinedReport += 1

                Case 5
                    Me.Controls.Add(dgvTemp5)
                    With dgvTemp5
                        .DataSource = Nothing
                        .DataSource = xlDataset
                        .DataMember = "[Sheet1$]"
                        .Visible = False
                        .AllowUserToAddRows = False
                    End With
                    Me.Controls.Add(dgvTemp5)
                    intCombinedReport += 1

                Case 6
                    Me.Controls.Add(dgvTemp6)
                    With dgvTemp6
                        .DataSource = Nothing
                        .DataSource = xlDataset
                        .DataMember = "[Sheet1$]"
                        .Visible = False
                        .AllowUserToAddRows = False
                    End With
                    intCombinedReport += 1

            End Select
            lblCombineReport.Text = "Combine Report: " & intCombinedReport

        End If

        Me.Cursor = Cursors.WaitCursor
        With dgvPreviewExcel
            For intColumn = 0 To .Columns.Count - 1
                DoEvents()
                .Columns(intColumn).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            Next

            'Remove column 0 if it is a checkbox
            'dgvPreviewExcel.Columns.RemoveAt(0)

            lblProcessing.Text = "Status: Please wait while checking report format..."


            If chkProductExpirationReport.Checked = False Then 'NOT Expired Product Report
                CheckHardwareReportFormat()

                If isCSVFormat = False Then
                    If blnHardwareReport = True And blnHardwareReportWCheckBox = True Then
                        'dgvPreviewExcel.Columns.RemoveAt(0) 'temporary disable
                    Else
                        .Columns(25).HeaderText = "Date Of Last Receipt"
                        .Columns(26).HeaderText = "Date Of Last Sale"
                    End If
                End If

                'Skip for High Permformance Report 7-7-21
                If blnHighPerformanceReport = False Then
                    ProcessMonthHeaders()
                End If



            End If
        End With

        lblColumns.Text = "Columns: " & dgvPreviewExcel.Columns.Count
        lblRows.Text = "Rows: " & dgvPreviewExcel.Rows.Count

        xlConnection.Close()
        lblProcessing.Text = "Status: Ready to process"
        butProcess.Enabled = True
        If blnCombineReport = True Then
            lsvFiles.Enabled = True
        End If
        Me.Cursor = Cursors.Default


        'Catch ex As Exception
        '    MessageBox.Show(ex.Message, "Error Occured", MessageBoxButtons.OK)
        'End Try

ExitTextProcess:
        'MessageBox.Show("Exit Text Process")
        'Me.Cursor = Cursors.Default
        'End

    End Sub

    Private Enum RegEntryName
        PrintAreaColumn1
        PrintAreaColumn2
        PrintPaperOrientation
        PrintPaperSize
    End Enum

    Private Sub AllowRegistryAccess()
        Dim CurrentUser As String = Environment.UserDomainName & "\" & Environment.UserName

        'RegSecurity.AddAccessRule(New RegistryAccessRule(CurrentUser, RegistryRights.ReadKey Or RegistryRights.Delete Or RegistryRights.SetValue Or RegistryRights.WriteKey Or RegistryRights.ChangePermissions, InheritanceFlags.None, _
        'PropagationFlags.None, AccessControlType.Allow))

    End Sub

    Private Sub RegistryWrite()
        Dim regKey As RegistryKey
        Dim regSec As RegistrySecurity
        Dim CurrentUser As String = Environment.UserDomainName & "\" & Environment.UserName

        'Try

        Me.Cursor = Cursors.WaitCursor
        regSec = New RegistrySecurity()
        regSec.AddAccessRule(New RegistryAccessRule(CurrentUser, RegistryRights.WriteKey, InheritanceFlags.None, PropagationFlags.None, AccessControlType.Allow))

        regKey = Registry.CurrentUser.OpenSubKey("SOFTWARE", True)
        regKey.SetAccessControl(regSec)

        regKey.CreateSubKey("EagleReportTool")

        regKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\EagleReportTool", True)
        regKey.SetAccessControl(regSec)

        With regKey
            .SetValue("_Version", "1.0")
            .SetValue("_Author", "Jeffrey Balbalosa")
            .SetValue("_Company", "Surangel and Sons, Co.")
            .SetValue("_Country", "Koror, Palau")

            .SetValue("ColorColumnHeader", "")
            .SetValue("ColorColumn1", "")
            .SetValue("ColorColumn2", "")
            .SetValue("ColorColumn3", "")
            .SetValue("ColorColumn4", "")
            .SetValue("ColorColumn5", "")
            .SetValue("ColorColumn6", "")
            .SetValue("MaxColumn", "")
            .SetValue("MultiReport", chkMultiReport.Checked)
            .SetValue("ClearOrderPoint", chkClearOrderPoint.Checked)
            .SetValue("PrintAreaColumn1", Trim(txtPrintAreaColumn1.Text))
            .SetValue("PrintAreaColumn2", Trim(txtPrintAreaColumn2.Text))
            .SetValue("PrintZoomPercentage", Trim(txtZoom.Text))
            .SetValue("PrintPaperOrientation", cboPaperOrientation.Text)
            .SetValue("PrintPaperSize", cboPaperSize.Text)
            .SetValue("ReportType", "")
            .SetValue("RangeStart", Trim(txtRangeStart.Text))
            .SetValue("RangeEnd", Trim(txtRangeEnd.Text))
            '.SetValue("SortColumnName", Trim(txtColumnSort.Text))
            .Close()
        End With

        Me.Cursor = Cursors.Default

        'Catch ex As Exception
        '    MessageBox.Show("Unable to write." & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        'End Try

    End Sub

    Private Sub RegistryWrite(ByVal KeyName As String, ByVal DataValue As String)
        Dim regKey As RegistryKey

        AllowRegistryAccess()
        regKey = Registry.CurrentUser.OpenSubKey("SOFTWARE", True)
        regKey.CreateSubKey("EagleReportTool", RegistryKeyPermissionCheck.Default)

        regKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\EagleReportTool", True)
        With regKey
            .SetValue(KeyName, DataValue)
            .Close()
        End With

    End Sub

    Private Sub RegistryRead()
        Dim RegKey As RegistryKey
        Dim RegSec As New RegistrySecurity
        Dim CurrentUser As String = Environment.UserDomainName & "\" & Environment.UserName
        'Dim res As Object

        RegSec.AddAccessRule(New RegistryAccessRule(CurrentUser, RegistryRights.ReadKey Or RegistryRights.Delete Or RegistryRights.SetValue Or RegistryRights.WriteKey Or RegistryRights.ChangePermissions,
                                                    InheritanceFlags.None, PropagationFlags.None, AccessControlType.Allow))

        'Try

        RegKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\EagleReportTool", False)

        With RegKey
            chkMultiReport.Checked = .GetValue("MultiReport")
            txtPrintAreaColumn1.Text = UCase(.GetValue("PrintAreaColumn1"))
            txtPrintAreaColumn2.Text = UCase(.GetValue("PrintAreaColumn2"))
            txtZoom.Text = .GetValue("PrintZoomPercentage")

            With cboPaperOrientation
                .Text = RegKey.GetValue("PrintPaperOrientation")
                If .Text = "" Then
                    .SelectedIndex = 0
                End If
            End With

            With cboPaperSize
                .Text = RegKey.GetValue("PrintPaperSize")
                If .Text = "" Then
                    .SelectedIndex = 0
                End If
            End With

            txtRangeStart.Text = UCase(.GetValue("RangeStart"))
            txtRangeEnd.Text = UCase(.GetValue("RangeEnd"))
            'txtColumnSort.Text = .GetValue("SortColumnName")
            txtSortLevel1.Text = UCase(.GetValue("SortLevel1"))
            txtSortLevel2.Text = UCase(.GetValue("SortLevel2"))
            txtSortLevel3.Text = UCase(.GetValue("SortLevel3"))

            strRangeStart = UCase(.GetValue("RangeStart"))
            strRangeEnd = UCase(.GetValue("RangeEnd"))

            strMultiReport = .GetValue("MultiReport")
            blnCombineReport = .GetValue("CombineReport")
            blnClearOrderPoint = .GetValue("ClearOrderPoint")
            blnProductExpirationReport = .GetValue("ProductExpirationReport")

            chkMultiReport.Checked = IIf(strMultiReport = "True", True, False)
            chkCombineReport.Checked = IIf(blnCombineReport = True, True, False)
            chkClearOrderPoint.Checked = IIf(blnClearOrderPoint = True, True, False)
            chkProductExpirationReport.Checked = IIf(blnProductExpirationReport = True, True, False)
            .Close()

        End With

        'Catch ex As Exception
        '    res = MessageBox.Show(ex.Message, "Error", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error)
        '    If res = vbAbort Then
        '        End
        '    End If
        'End Try

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing

        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub LoadFiles(ByVal FileDirectory As String)

        Dim dDirectories() As String = IO.Directory.GetDirectories(FileDirectory)
        Dim strDirectory As String = ""

        Me.Cursor = Cursors.WaitCursor

        Dim strItem(4) As String
        Dim lsvItem As ListViewItem

        Dim Files() As String = IO.Directory.GetFiles(FileDirectory)

        'Load each files in a directory
        For Each File As String In Files
            If Mid(Path.GetFileName(File.ToString), 2, 1) <> "$" And (Path.GetExtension(File.ToString) = ".xlsx" Or Path.GetExtension(File.ToString) = ".xls" Or Path.GetExtension(File.ToString) = ".csv") Then
                strItem(0) = Path.GetFileName(File.ToString)
                strItem(1) = IO.File.GetCreationTime(File.ToString)
                strItem(2) = FileDirectory

                lsvItem = New ListViewItem(strItem)
                lsvFiles.Items.Add(lsvItem)
                lsvFiles.Sorting = System.Windows.Forms.SortOrder.Ascending
                lsvFiles.Sort()
            End If
        Next




        'lblLoadDirectory.Text = lsvFiles.Items.Count
        'lblProcessing.Text = "Status: Ready"
        'Me.Cursor = Cursors.Default

        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


        'Dim dDirectories() As String = IO.Directory.GetDirectories(FileDirectory)
        'Dim strDirectory As String = ""

        If dDirectories.Length <> 0 Then

            For Each Dir As String In dDirectories
                DoEvents()
                Dim dFiles As New IO.DirectoryInfo(Dir.ToString)
                Dim dGetFiles As IO.FileInfo() = dFiles.GetFiles()
                Dim dFile As IO.FileInfo
                'Dim strItem(3) As String
                'Dim lsvItem As ListViewItem

                For Each dFile In dGetFiles
                    DoEvents()
                    If Mid(dFile.ToString, 1, 2) <> "~$" And (Path.GetExtension(dFile.ToString) = ".xlsx" Or Path.GetExtension(dFile.ToString) = ".xls" Or Path.GetExtension(dFile.ToString) = ".csv") Then
                        strItem(0) = Path.GetFileName(dFile.ToString)
                        strItem(1) = dFile.CreationTime    'File.GetCreationTime(dFile.ToString)
                        strItem(2) = Dir.ToString & "\" & dFile.ToString

                        lsvItem = New ListViewItem(strItem)
                        lsvFiles.Items.Add(lsvItem)
                        lsvFiles.Sorting = System.Windows.Forms.SortOrder.Ascending
                        lsvFiles.Sort()
                    End If
                Next

                dFiles = Nothing
                dGetFiles = Nothing
            Next

        Else
            Dim strFiles As String
            'Dim strItem(3) As String
            'Dim lsvItem As ListViewItem

            For Each strFiles In Directory.GetFiles(FileDirectory)
                DoEvents()
                If Mid(strFiles.ToString, 1, 2) <> "~$" And (Path.GetExtension(strFiles.ToString) = ".xlsx" Or Path.GetExtension(strFiles.ToString) = ".xls" Or Path.GetExtension(strFiles.ToString) = ".csv") Then
                    strItem(0) = Path.GetFileName(strFiles.ToString)
                    strItem(1) = IO.File.GetLastWriteTime(strFiles.ToString)
                    strItem(2) = strFiles.ToString
                    lsvItem = New ListViewItem(strItem)
                    lsvFiles.Items.Add(lsvItem)
                    lsvFiles.Sorting = System.Windows.Forms.SortOrder.Ascending
                    lsvFiles.Sort()

                End If
            Next

        End If


        lblProcessing.Text = ""

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub CheckSurangelFolder()
        If Directory.Exists(DirPathEpicor & "\Surangel") = False Then
            Directory.CreateDirectory(DirPathEpicor & "\Surangel")
        End If
    End Sub

    Private Function CharPrior(ByVal ColumnNumber As Integer, Optional ByVal CharToAdd As String = "'") As String
        Dim strChar As String = ""

        Select Case ColumnNumber
            Case 0 : strChar = CharToAdd
            Case 1 : strChar = CharToAdd
            Case 2 : strChar = CharToAdd
            Case 3 : strChar = CharToAdd
                'Case 5 : strChar = CharToAdd
            Case 35 : strChar = CharToAdd
        End Select

        Return strChar

    End Function



    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RegistryRead()

        Me.Text = "Surangel and Sons - Eagle Report Tool"
        lblProcessing.Text = "Status: "

        If Directory.Exists(DirPathEpicor) = False Then
            Directory.CreateDirectory("C:\Users\" & Environment.UserName & "\AppData\Local\Temp\Epicor")
        End If

        If Directory.Exists(DirPath3Apps) = False Then
            MessageBox.Show("Eagle 3apps file directory doesn't exist. Please contact your admin", "Directory not found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End
        End If

        LoadFiles(DirPathEpicor)
        LoadFiles(DirPath3Apps)
    End Sub

    Private Sub butClose_Click(sender As Object, e As EventArgs) Handles butClose.Click
        RegistryWrite()
        End

    End Sub

    Private Sub butRefresh_Click(sender As Object, e As EventArgs) Handles butRefresh.Click
        Me.Cursor = Cursors.WaitCursor
        lsvFiles.Enabled = True
        lsvFiles.Items.Clear()

        LoadFiles(DirPathEpicor)
        LoadFiles(DirPath3Apps)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub butClear_Click(sender As Object, e As EventArgs) Handles butClear.Click
        blnHighPerformanceReport = False
        dgvPreviewExcel.DataSource = Nothing
        lblRows.Text = "Rows:"
        lblColumns.Text = "Columns:"
        lsvFiles.Enabled = True
        butProcess.Enabled = False
    End Sub

    Private Sub txtSortLevel1_TextChanged(sender As Object, e As EventArgs) Handles txtSortLevel1.TextChanged
        If txtSortLevel1.Text <> "" Then
            RegistryWrite("SortLevel1", Trim(txtSortLevel1.Text))
        End If
    End Sub

    Private Sub txtSortLevel2_TextChanged(sender As Object, e As EventArgs) Handles txtSortLevel2.TextChanged
        If txtSortLevel2.Text <> "" Then
            RegistryWrite("SortLevel2", Trim(txtSortLevel2.Text))
        End If
    End Sub

    Private Sub txtSortLevel3_TextChanged(sender As Object, e As EventArgs) Handles txtSortLevel3.TextChanged
        If txtSortLevel3.Text <> "" Then
            RegistryWrite("SortLevel3", Trim(txtSortLevel3.Text))
        End If
    End Sub

    Private Sub lsvFiles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lsvFiles.SelectedIndexChanged
        Dim intLoop As Integer = 0

        Me.Cursor = Cursors.WaitCursor
        blnHighPerformanceReport = False
        butProcess.Enabled = False
        If chkCombineReport.Checked = False Then
            lsvFiles.Enabled = False
            blnFirstReportProcessed = True
        End If

        For intLoop = 0 To lsvFiles.Items.Count - 1
            DoEvents()
            If lsvFiles.Items(intLoop).Selected = True Then
                strFileDirectory = lsvFiles.Items(intLoop).SubItems(2).Text
                Preview(lsvFiles.Items(intLoop).Text, strFileDirectory)
                butProcess.Enabled = IIf(blnCombineReport = True, False, True)


                If blnHardwareReport = False Then
                    'ComputeAvailableQuantity()
                End If


                intDGVPreview += 1
                Exit For

            End If
        Next
        If intDGVPreview > 1 Then
            butProcess.Enabled = True
        End If

        Me.Cursor = Cursors.Default

        'butProcess_Click(sender, e)
    End Sub

    Private Sub butProcess_Click(sender As Object, e As EventArgs) Handles butProcess.Click
        Me.Cursor = Cursors.WaitCursor

        CheckSurangelFolder()

        If chkProductExpirationReport.Checked = True Then 'Expired Report
            Dim xlApp As New Excel.Application
            Dim xlConnection As New OleDb.OleDbConnection


            xlConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.14.0; Data Source= " & strExcelFilePathExpiredProduct & "\" & strExcelFileNameExpiredProduct & ";Extended Properties = 'Excel 12.0 Xml;HDR=YES;'")
            xlConnection.Open()

            Dim strSKU As String = ""
            Dim strUPC As String = ""
            Dim strQty As String = ""
            Dim strExpirationDate As String = ""


            Dim TxtReader As TextReader = New StreamReader(strTextFilePathExpiredProduct & "\" & strTextFileNameExpiredProduct)
            Dim recordLine As String = ""
            Dim iRecordLine As Integer = 0
            Dim iLoop As Integer = 0
            Dim iQOHTotal As Integer = 0

            Do Until recordLine Is Nothing
                DoEvents()
                iRecordLine += 1

                Dim xlDataset As New System.Data.DataSet
                Dim xlDataAdapter As New System.Data.OleDb.OleDbDataAdapter
                Dim xlCommand As New System.Data.OleDb.OleDbCommand

                xlCommand.Connection = xlConnection
                xlCommand.CommandType = CommandType.Text

                If TxtReader.Peek = -1 Then
                    Exit Do
                End If

                recordLine = TxtReader.ReadLine()
                Dim list As IList(Of String) = New List(Of String)(recordLine.Split(New String() {","}, StringSplitOptions.None))

                strSKU = Trim(list(0))
                strQty = Trim(list(1))
                strExpirationDate = Trim(list(2))

                lblProcessing.Text = "Processing Item: " & strSKU & " Line No: " & iRecordLine
                xlCommand.CommandText = "select * from [Sheet1$] where [Item Number]='" & strSKU & "'"
                xlCommand.ExecuteNonQuery()
                xlDataAdapter.SelectCommand = xlCommand
                xlDataAdapter.TableMappings.Add("Table", "Sheet1")
                xlDataAdapter.Fill(xlDataset)


                If xlDataset.Tables(0).Rows.Count > 0 Then
                    With dgvPreviewExcel
                        .Rows.Add()
                        .Rows(.RowCount - 1).Cells("SKU").Value = strSKU
                        .Rows(.RowCount - 1).Cells("ExpirationDate").Value = strExpirationDate

                        If xlDataset.Tables(0).Rows(0).Item("Item Description").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("ItemDescription").Value = Trim(xlDataset.Tables(0).Rows(0).Item("Item Description"))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item("Pack").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("Pack").Value = Trim(xlDataset.Tables(0).Rows(0).Item("Pack"))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item("PUR UOM").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("PURUOM").Value = Trim(xlDataset.Tables(0).Rows(0).Item("PUR UOM"))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item("UPC Code").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("UPC").Value = Trim(xlDataset.Tables(0).Rows(0).Item("UPC Code"))
                        End If

                        iQOHTotal = 0
                        For iLoop = 0 To xlDataset.Tables(0).Rows.Count - 1
                            DoEvents()
                            If xlDataset.Tables(0).Rows(iLoop).Item("QOH").ToString <> "" Then
                                iQOHTotal += Convert.ToInt16(xlDataset.Tables(0).Rows(iLoop).Item("QOH"))
                                Select Case iLoop
                                    Case 0 : .Rows(.RowCount - 1).Cells("QOHST1").Value = Trim(xlDataset.Tables(0).Rows(iLoop).Item("QOH"))
                                    Case 1 : .Rows(.RowCount - 1).Cells("QOHST3").Value = Trim(xlDataset.Tables(0).Rows(iLoop).Item("QOH"))
                                    Case 2 : .Rows(.RowCount - 1).Cells("QOHST7").Value = Trim(xlDataset.Tables(0).Rows(iLoop).Item("QOH"))
                                End Select
                            End If
                        Next

                        .Rows(.RowCount - 1).Cells("QOHTotal").Value = iQOHTotal

                        If xlDataset.Tables(0).Rows(0).Item("Order Point").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("OrderPoint").Value = Trim(xlDataset.Tables(0).Rows(0).Item("Order Point"))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item("QOO").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("QOO").Value = Trim(xlDataset.Tables(0).Rows(0).Item("QOO"))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item(9).ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("SalesUnits1").Value = Trim(xlDataset.Tables(0).Rows(0).Item(9))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item(10).ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("SalesUnits2").Value = Trim(xlDataset.Tables(0).Rows(0).Item(10))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item(11).ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("SalesUnits3").Value = Trim(xlDataset.Tables(0).Rows(0).Item(11))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item(12).ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("SalesUnits4").Value = Trim(xlDataset.Tables(0).Rows(0).Item(12))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item("Retail Price").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("RetailPrice").Value = Trim(xlDataset.Tables(0).Rows(0).Item("Retail Price"))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item("Average Cost").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("AvgCost").Value = Trim(xlDataset.Tables(0).Rows(0).Item("Average Cost"))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item("GMROI").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("GMROI").Value = Trim(xlDataset.Tables(0).Rows(0).Item("GMROI"))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item("Repl Cost").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("ReplCost").Value = Trim(xlDataset.Tables(0).Rows(0).Item("Repl Cost"))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item("Date Of Last Sale").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("DateOfLastSale").Value = FormatDateTime(xlDataset.Tables(0).Rows(0).Item("Date Of Last Sale"), DateFormat.ShortDate)
                        End If

                        If xlDataset.Tables(0).Rows(0).Item("Date Of Last Receipt").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("DateOfLastReceipt").Value = FormatDateTime(xlDataset.Tables(0).Rows(0).Item("Date Of Last Receipt"), DateFormat.ShortDate)
                        End If

                        If xlDataset.Tables(0).Rows(0).Item("Department Code").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("DepartmentCode").Value = Trim(xlDataset.Tables(0).Rows(0).Item("Department Code"))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item(21).ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("MFGPart#").Value = Trim(xlDataset.Tables(0).Rows(0).Item(21))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item("Sales Units YTD").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("SalesUnitYTD").Value = Trim(xlDataset.Tables(0).Rows(0).Item("Sales Units YTD"))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item("Last Year Units").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("LastYearUnits").Value = Trim(xlDataset.Tables(0).Rows(0).Item("Last Year Units"))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item("Retail Price Change").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("RetailPriceChange").Value = Trim(xlDataset.Tables(0).Rows(0).Item("Retail Price Change"))
                        End If

                        If xlDataset.Tables(0).Rows(0).Item("Vendor Name").ToString <> "" Then
                            .Rows(.RowCount - 1).Cells("VendorName").Value = Trim(xlDataset.Tables(0).Rows(0).Item("Vendor Name"))
                        End If

                    End With
                End If

                xlCommand = Nothing
                xlDataAdapter = Nothing
                xlDataset = Nothing

            Loop

            xlConnection.Close()
            ExportExpiredReportToExcel(dgvPreviewExcel)


        ElseIf blnHardwareReport = False Then
            ExportDataToExcel(dgvPreviewExcel)

        ElseIf blnHardwareReport = True Then
            chkMultiReport.Checked = False
            chkPOImport.Checked = False

            If blnCombineReport = False Then
                ExportHardwareReport(dgvPreviewExcel)
            Else
                With dgvPreviewExcel
                    .DataSource = Nothing

                    Dim intCol As Integer = 0
                    Dim intRow As Integer = 0
                    For intCol = 0 To dgvTemp0.ColumnCount - 1
                        .Columns.Add(dgvTemp0.Columns(intCol).Name, dgvTemp0.Columns(intCol).Name)
                    Next

                    If dgvTemp0.RowCount <> 0 Then
                        For intRow = 0 To dgvTemp0.RowCount - 1
                            .Rows.Add()
                            For intCol = 0 To dgvTemp0.ColumnCount - 1
                                .Rows(.RowCount - 1).Cells(intCol).Value = dgvTemp0.Rows(intRow).Cells(intCol).Value
                            Next
                        Next
                    End If

                    If dgvTemp1.RowCount <> 0 Then
                        For intRow = 0 To dgvTemp1.RowCount - 1
                            .Rows.Add()
                            For intCol = 0 To dgvTemp1.ColumnCount - 1
                                .Rows(.RowCount - 1).Cells(intCol).Value = dgvTemp1.Rows(intRow).Cells(intCol).Value
                            Next
                        Next
                    End If

                    If dgvTemp2.RowCount <> 0 Then
                        For intRow = 0 To dgvTemp0.RowCount - 1
                            .Rows.Add()
                            For intCol = 0 To dgvTemp2.ColumnCount - 1
                                .Rows(.RowCount - 1).Cells(intCol).Value = dgvTemp2.Rows(intRow).Cells(intCol).Value
                            Next
                        Next
                    End If

                    If dgvTemp3.RowCount <> 0 Then
                        For intRow = 0 To dgvTemp3.RowCount - 1
                            .Rows.Add()
                            For intCol = 0 To dgvTemp3.ColumnCount - 1
                                .Rows(.RowCount - 1).Cells(intCol).Value = dgvTemp3.Rows(intRow).Cells(intCol).Value
                            Next
                        Next
                    End If

                    If dgvTemp4.RowCount <> 0 Then
                        For intRow = 0 To dgvTemp4.RowCount - 1
                            .Rows.Add()
                            For intCol = 0 To dgvTemp4.ColumnCount - 1
                                .Rows(.RowCount - 1).Cells(intCol).Value = dgvTemp4.Rows(intRow).Cells(intCol).Value
                            Next
                        Next
                    End If

                    If dgvTemp5.RowCount <> 0 Then
                        For intRow = 0 To dgvTemp5.RowCount - 1
                            .Rows.Add()
                            For intCol = 0 To dgvTemp5.ColumnCount - 1
                                .Rows(.RowCount - 1).Cells(intCol).Value = dgvTemp5.Rows(intRow).Cells(intCol).Value
                            Next
                        Next
                    End If

                    If dgvTemp6.RowCount <> 0 Then
                        For intRow = 0 To dgvTemp6.RowCount - 1
                            .Rows.Add()
                            For intCol = 0 To dgvTemp6.ColumnCount - 1
                                .Rows(.RowCount - 1).Cells(intCol).Value = dgvTemp6.Rows(intRow).Cells(intCol).Value
                            Next
                        Next
                    End If

                End With

                ExportHardwareReport(dgvPreviewExcel)

            End If
        End If

        tpbExport.Value = 0
        lblProcessing.Text = "Status: "
        lsvFiles.Enabled = True

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub txtPrintAreaColumn1_LostFocus(sender As Object, e As EventArgs) Handles txtPrintAreaColumn1.LostFocus
        txtPrintAreaColumn1.Text = txtPrintAreaColumn1.Text.ToUpper()

    End Sub
End Class
