<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.lsvFiles = New System.Windows.Forms.ListView()
        Me.lsvName = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.lsvDateModified = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.lsvFilePath = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.dgvPreviewExcel = New System.Windows.Forms.DataGridView()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.butClear = New System.Windows.Forms.ToolStripButton()
        Me.butRefresh = New System.Windows.Forms.ToolStripButton()
        Me.butProcess = New System.Windows.Forms.ToolStripButton()
        Me.butOption = New System.Windows.Forms.ToolStripButton()
        Me.butClose = New System.Windows.Forms.ToolStripButton()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblRows = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblColumns = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblProcessing = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tpbExport = New System.Windows.Forms.ToolStripProgressBar()
        Me.lblCombineReport = New System.Windows.Forms.ToolStripStatusLabel()
        Me.grbReportOption = New System.Windows.Forms.GroupBox()
        Me.chkProductExpirationReport = New System.Windows.Forms.CheckBox()
        Me.chkExtendedReportZ = New System.Windows.Forms.CheckBox()
        Me.chkClearOrderPoint = New System.Windows.Forms.CheckBox()
        Me.chkCombineReport = New System.Windows.Forms.CheckBox()
        Me.chkPOImport = New System.Windows.Forms.CheckBox()
        Me.chkMultiReport = New System.Windows.Forms.CheckBox()
        Me.grbPrintSetup = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboPaperSize = New System.Windows.Forms.ComboBox()
        Me.cboPaperOrientation = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtZoom = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtPrintAreaColumn2 = New System.Windows.Forms.TextBox()
        Me.txtPrintAreaColumn1 = New System.Windows.Forms.TextBox()
        Me.grbWorkSheet = New System.Windows.Forms.GroupBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtSortLevel3 = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtSortLevel2 = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtSortLevel1 = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtRangeEnd = New System.Windows.Forms.TextBox()
        Me.txtRangeStart = New System.Windows.Forms.TextBox()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvPreviewExcel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.grbReportOption.SuspendLayout()
        Me.grbPrintSetup.SuspendLayout()
        Me.grbWorkSheet.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(12, 55)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.lsvFiles)
        Me.SplitContainer1.Panel1.RightToLeft = System.Windows.Forms.RightToLeft.No
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvPreviewExcel)
        Me.SplitContainer1.Panel2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SplitContainer1.Size = New System.Drawing.Size(1002, 636)
        Me.SplitContainer1.SplitterDistance = 290
        Me.SplitContainer1.TabIndex = 0
        '
        'lsvFiles
        '
        Me.lsvFiles.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lsvFiles.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lsvFiles.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.lsvName, Me.lsvDateModified, Me.lsvFilePath})
        Me.lsvFiles.HideSelection = False
        Me.lsvFiles.Location = New System.Drawing.Point(3, 3)
        Me.lsvFiles.Name = "lsvFiles"
        Me.lsvFiles.Size = New System.Drawing.Size(290, 626)
        Me.lsvFiles.TabIndex = 2
        Me.lsvFiles.UseCompatibleStateImageBehavior = False
        Me.lsvFiles.View = System.Windows.Forms.View.Details
        '
        'lsvName
        '
        Me.lsvName.Text = "Name"
        Me.lsvName.Width = 140
        '
        'lsvDateModified
        '
        Me.lsvDateModified.Text = "Date Modified"
        Me.lsvDateModified.Width = 100
        '
        'lsvFilePath
        '
        Me.lsvFilePath.Text = "File Path"
        '
        'dgvPreviewExcel
        '
        Me.dgvPreviewExcel.AllowUserToAddRows = False
        Me.dgvPreviewExcel.AllowUserToDeleteRows = False
        Me.dgvPreviewExcel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvPreviewExcel.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgvPreviewExcel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPreviewExcel.Location = New System.Drawing.Point(3, 3)
        Me.dgvPreviewExcel.MultiSelect = False
        Me.dgvPreviewExcel.Name = "dgvPreviewExcel"
        Me.dgvPreviewExcel.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvPreviewExcel.Size = New System.Drawing.Size(698, 631)
        Me.dgvPreviewExcel.StandardTab = True
        Me.dgvPreviewExcel.TabIndex = 0
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.butClear, Me.butRefresh, Me.butProcess, Me.butOption, Me.butClose})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1223, 52)
        Me.ToolStrip1.TabIndex = 1
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'butClear
        '
        Me.butClear.Image = CType(resources.GetObject("butClear.Image"), System.Drawing.Image)
        Me.butClear.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.butClear.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butClear.Name = "butClear"
        Me.butClear.Size = New System.Drawing.Size(38, 49)
        Me.butClear.Text = "Clear"
        Me.butClear.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'butRefresh
        '
        Me.butRefresh.Image = CType(resources.GetObject("butRefresh.Image"), System.Drawing.Image)
        Me.butRefresh.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.butRefresh.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butRefresh.Name = "butRefresh"
        Me.butRefresh.Size = New System.Drawing.Size(50, 49)
        Me.butRefresh.Text = "Refresh"
        Me.butRefresh.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'butProcess
        '
        Me.butProcess.Image = CType(resources.GetObject("butProcess.Image"), System.Drawing.Image)
        Me.butProcess.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.butProcess.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butProcess.Name = "butProcess"
        Me.butProcess.Size = New System.Drawing.Size(51, 49)
        Me.butProcess.Text = "Process"
        Me.butProcess.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'butOption
        '
        Me.butOption.Image = CType(resources.GetObject("butOption.Image"), System.Drawing.Image)
        Me.butOption.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.butOption.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butOption.Name = "butOption"
        Me.butOption.Size = New System.Drawing.Size(48, 49)
        Me.butOption.Text = "Option"
        Me.butOption.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'butClose
        '
        Me.butClose.Image = CType(resources.GetObject("butClose.Image"), System.Drawing.Image)
        Me.butClose.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.butClose.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butClose.Name = "butClose"
        Me.butClose.Size = New System.Drawing.Size(40, 49)
        Me.butClose.Text = "Close"
        Me.butClose.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel2, Me.lblRows, Me.lblColumns, Me.lblProcessing, Me.tpbExport, Me.lblCombineReport})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 694)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1223, 22)
        Me.StatusStrip1.TabIndex = 2
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(63, 17)
        Me.ToolStripStatusLabel1.Text = "Version 4.0"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(117, 17)
        Me.ToolStripStatusLabel2.Text = "Published: 8/26/2021"
        '
        'lblRows
        '
        Me.lblRows.Name = "lblRows"
        Me.lblRows.Size = New System.Drawing.Size(38, 17)
        Me.lblRows.Text = "Rows:"
        '
        'lblColumns
        '
        Me.lblColumns.Name = "lblColumns"
        Me.lblColumns.Size = New System.Drawing.Size(58, 17)
        Me.lblColumns.Text = "Columns:"
        '
        'lblProcessing
        '
        Me.lblProcessing.Name = "lblProcessing"
        Me.lblProcessing.Size = New System.Drawing.Size(64, 17)
        Me.lblProcessing.Text = "Processing"
        '
        'tpbExport
        '
        Me.tpbExport.Name = "tpbExport"
        Me.tpbExport.Size = New System.Drawing.Size(100, 16)
        Me.tpbExport.Visible = False
        '
        'lblCombineReport
        '
        Me.lblCombineReport.Name = "lblCombineReport"
        Me.lblCombineReport.Size = New System.Drawing.Size(94, 17)
        Me.lblCombineReport.Text = "Combine Report"
        Me.lblCombineReport.Visible = False
        '
        'grbReportOption
        '
        Me.grbReportOption.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grbReportOption.Controls.Add(Me.chkProductExpirationReport)
        Me.grbReportOption.Controls.Add(Me.chkExtendedReportZ)
        Me.grbReportOption.Controls.Add(Me.chkClearOrderPoint)
        Me.grbReportOption.Controls.Add(Me.chkCombineReport)
        Me.grbReportOption.Controls.Add(Me.chkPOImport)
        Me.grbReportOption.Controls.Add(Me.chkMultiReport)
        Me.grbReportOption.Location = New System.Drawing.Point(1020, 44)
        Me.grbReportOption.Name = "grbReportOption"
        Me.grbReportOption.Size = New System.Drawing.Size(191, 185)
        Me.grbReportOption.TabIndex = 3
        Me.grbReportOption.TabStop = False
        Me.grbReportOption.Text = "Report Option"
        '
        'chkProductExpirationReport
        '
        Me.chkProductExpirationReport.AutoSize = True
        Me.chkProductExpirationReport.Location = New System.Drawing.Point(22, 147)
        Me.chkProductExpirationReport.Name = "chkProductExpirationReport"
        Me.chkProductExpirationReport.Size = New System.Drawing.Size(112, 17)
        Me.chkProductExpirationReport.TabIndex = 5
        Me.chkProductExpirationReport.Text = "Product Expiration"
        Me.chkProductExpirationReport.UseVisualStyleBackColor = True
        '
        'chkExtendedReportZ
        '
        Me.chkExtendedReportZ.AutoSize = True
        Me.chkExtendedReportZ.Location = New System.Drawing.Point(22, 124)
        Me.chkExtendedReportZ.Name = "chkExtendedReportZ"
        Me.chkExtendedReportZ.Size = New System.Drawing.Size(116, 17)
        Me.chkExtendedReportZ.TabIndex = 4
        Me.chkExtendedReportZ.Text = "Extended Report Z"
        Me.chkExtendedReportZ.UseVisualStyleBackColor = True
        '
        'chkClearOrderPoint
        '
        Me.chkClearOrderPoint.AutoSize = True
        Me.chkClearOrderPoint.Checked = True
        Me.chkClearOrderPoint.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkClearOrderPoint.Location = New System.Drawing.Point(22, 101)
        Me.chkClearOrderPoint.Name = "chkClearOrderPoint"
        Me.chkClearOrderPoint.Size = New System.Drawing.Size(106, 17)
        Me.chkClearOrderPoint.TabIndex = 3
        Me.chkClearOrderPoint.Text = "Clear Order Point"
        Me.chkClearOrderPoint.UseVisualStyleBackColor = True
        '
        'chkCombineReport
        '
        Me.chkCombineReport.AutoSize = True
        Me.chkCombineReport.Location = New System.Drawing.Point(22, 78)
        Me.chkCombineReport.Name = "chkCombineReport"
        Me.chkCombineReport.Size = New System.Drawing.Size(102, 17)
        Me.chkCombineReport.TabIndex = 2
        Me.chkCombineReport.Text = "Combine Report"
        Me.chkCombineReport.UseVisualStyleBackColor = True
        '
        'chkPOImport
        '
        Me.chkPOImport.AutoSize = True
        Me.chkPOImport.Location = New System.Drawing.Point(22, 55)
        Me.chkPOImport.Name = "chkPOImport"
        Me.chkPOImport.Size = New System.Drawing.Size(73, 17)
        Me.chkPOImport.TabIndex = 1
        Me.chkPOImport.Text = "PO Import"
        Me.chkPOImport.UseVisualStyleBackColor = True
        '
        'chkMultiReport
        '
        Me.chkMultiReport.AutoSize = True
        Me.chkMultiReport.Location = New System.Drawing.Point(22, 32)
        Me.chkMultiReport.Name = "chkMultiReport"
        Me.chkMultiReport.Size = New System.Drawing.Size(83, 17)
        Me.chkMultiReport.TabIndex = 0
        Me.chkMultiReport.Text = "Multi Report"
        Me.chkMultiReport.UseVisualStyleBackColor = True
        '
        'grbPrintSetup
        '
        Me.grbPrintSetup.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grbPrintSetup.Controls.Add(Me.Label5)
        Me.grbPrintSetup.Controls.Add(Me.Label4)
        Me.grbPrintSetup.Controls.Add(Me.cboPaperSize)
        Me.grbPrintSetup.Controls.Add(Me.cboPaperOrientation)
        Me.grbPrintSetup.Controls.Add(Me.Label3)
        Me.grbPrintSetup.Controls.Add(Me.txtZoom)
        Me.grbPrintSetup.Controls.Add(Me.Label2)
        Me.grbPrintSetup.Controls.Add(Me.Label1)
        Me.grbPrintSetup.Controls.Add(Me.txtPrintAreaColumn2)
        Me.grbPrintSetup.Controls.Add(Me.txtPrintAreaColumn1)
        Me.grbPrintSetup.Location = New System.Drawing.Point(1020, 235)
        Me.grbPrintSetup.Name = "grbPrintSetup"
        Me.grbPrintSetup.Size = New System.Drawing.Size(191, 169)
        Me.grbPrintSetup.TabIndex = 4
        Me.grbPrintSetup.TabStop = False
        Me.grbPrintSetup.Text = "Print Setup"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(19, 135)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(27, 13)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Size"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(19, 108)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(58, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Orientation"
        '
        'cboPaperSize
        '
        Me.cboPaperSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPaperSize.FormattingEnabled = True
        Me.cboPaperSize.Items.AddRange(New Object() {"Legal", "Letter"})
        Me.cboPaperSize.Location = New System.Drawing.Point(83, 135)
        Me.cboPaperSize.Name = "cboPaperSize"
        Me.cboPaperSize.Size = New System.Drawing.Size(91, 21)
        Me.cboPaperSize.TabIndex = 7
        '
        'cboPaperOrientation
        '
        Me.cboPaperOrientation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPaperOrientation.FormattingEnabled = True
        Me.cboPaperOrientation.Items.AddRange(New Object() {"Landscape", "Portrait"})
        Me.cboPaperOrientation.Location = New System.Drawing.Point(83, 108)
        Me.cboPaperOrientation.Name = "cboPaperOrientation"
        Me.cboPaperOrientation.Size = New System.Drawing.Size(91, 21)
        Me.cboPaperOrientation.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(19, 82)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Zoom %"
        '
        'txtZoom
        '
        Me.txtZoom.Location = New System.Drawing.Point(83, 82)
        Me.txtZoom.Name = "txtZoom"
        Me.txtZoom.Size = New System.Drawing.Size(91, 20)
        Me.txtZoom.TabIndex = 4
        Me.txtZoom.Text = "90"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(19, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Column 2"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(19, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Column 1"
        '
        'txtPrintAreaColumn2
        '
        Me.txtPrintAreaColumn2.Location = New System.Drawing.Point(83, 56)
        Me.txtPrintAreaColumn2.Name = "txtPrintAreaColumn2"
        Me.txtPrintAreaColumn2.Size = New System.Drawing.Size(91, 20)
        Me.txtPrintAreaColumn2.TabIndex = 1
        Me.txtPrintAreaColumn2.Text = "AA"
        '
        'txtPrintAreaColumn1
        '
        Me.txtPrintAreaColumn1.Location = New System.Drawing.Point(83, 27)
        Me.txtPrintAreaColumn1.Name = "txtPrintAreaColumn1"
        Me.txtPrintAreaColumn1.Size = New System.Drawing.Size(91, 20)
        Me.txtPrintAreaColumn1.TabIndex = 0
        Me.txtPrintAreaColumn1.Text = "B"
        '
        'grbWorkSheet
        '
        Me.grbWorkSheet.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grbWorkSheet.Controls.Add(Me.Label10)
        Me.grbWorkSheet.Controls.Add(Me.txtSortLevel3)
        Me.grbWorkSheet.Controls.Add(Me.Label9)
        Me.grbWorkSheet.Controls.Add(Me.txtSortLevel2)
        Me.grbWorkSheet.Controls.Add(Me.Label6)
        Me.grbWorkSheet.Controls.Add(Me.txtSortLevel1)
        Me.grbWorkSheet.Controls.Add(Me.Label7)
        Me.grbWorkSheet.Controls.Add(Me.Label8)
        Me.grbWorkSheet.Controls.Add(Me.txtRangeEnd)
        Me.grbWorkSheet.Controls.Add(Me.txtRangeStart)
        Me.grbWorkSheet.Location = New System.Drawing.Point(1020, 410)
        Me.grbWorkSheet.Name = "grbWorkSheet"
        Me.grbWorkSheet.Size = New System.Drawing.Size(191, 175)
        Me.grbWorkSheet.TabIndex = 7
        Me.grbWorkSheet.TabStop = False
        Me.grbWorkSheet.Text = "Worksheet"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(19, 140)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(64, 13)
        Me.Label10.TabIndex = 15
        Me.Label10.Text = "Sort Level 3"
        '
        'txtSortLevel3
        '
        Me.txtSortLevel3.Location = New System.Drawing.Point(83, 139)
        Me.txtSortLevel3.Name = "txtSortLevel3"
        Me.txtSortLevel3.Size = New System.Drawing.Size(91, 20)
        Me.txtSortLevel3.TabIndex = 14
        Me.txtSortLevel3.Text = "AM1"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(19, 114)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(64, 13)
        Me.Label9.TabIndex = 13
        Me.Label9.Text = "Sort Level 2"
        '
        'txtSortLevel2
        '
        Me.txtSortLevel2.Location = New System.Drawing.Point(83, 113)
        Me.txtSortLevel2.Name = "txtSortLevel2"
        Me.txtSortLevel2.Size = New System.Drawing.Size(91, 20)
        Me.txtSortLevel2.TabIndex = 12
        Me.txtSortLevel2.Text = "AL1"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(19, 88)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 13)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Sort Level 1"
        '
        'txtSortLevel1
        '
        Me.txtSortLevel1.Location = New System.Drawing.Point(83, 87)
        Me.txtSortLevel1.Name = "txtSortLevel1"
        Me.txtSortLevel1.Size = New System.Drawing.Size(91, 20)
        Me.txtSortLevel1.TabIndex = 10
        Me.txtSortLevel1.Text = "B1"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(19, 61)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(61, 13)
        Me.Label7.TabIndex = 9
        Me.Label7.Text = "End Range"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(19, 32)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(64, 13)
        Me.Label8.TabIndex = 8
        Me.Label8.Text = "Start Range"
        '
        'txtRangeEnd
        '
        Me.txtRangeEnd.Location = New System.Drawing.Point(83, 62)
        Me.txtRangeEnd.Name = "txtRangeEnd"
        Me.txtRangeEnd.Size = New System.Drawing.Size(91, 20)
        Me.txtRangeEnd.TabIndex = 7
        Me.txtRangeEnd.Text = "AW"
        '
        'txtRangeStart
        '
        Me.txtRangeStart.Location = New System.Drawing.Point(83, 33)
        Me.txtRangeStart.Name = "txtRangeStart"
        Me.txtRangeStart.Size = New System.Drawing.Size(91, 20)
        Me.txtRangeStart.TabIndex = 6
        Me.txtRangeStart.Text = "A"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1223, 716)
        Me.Controls.Add(Me.grbWorkSheet)
        Me.Controls.Add(Me.grbPrintSetup)
        Me.Controls.Add(Me.grbReportOption)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMain"
        Me.Text = "Eagle Report Tool"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvPreviewExcel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.grbReportOption.ResumeLayout(False)
        Me.grbReportOption.PerformLayout()
        Me.grbPrintSetup.ResumeLayout(False)
        Me.grbPrintSetup.PerformLayout()
        Me.grbWorkSheet.ResumeLayout(False)
        Me.grbWorkSheet.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents butClear As ToolStripButton
    Friend WithEvents butRefresh As ToolStripButton
    Friend WithEvents butProcess As ToolStripButton
    Friend WithEvents butOption As ToolStripButton
    Friend WithEvents butClose As ToolStripButton
    Friend WithEvents lsvFiles As ListView
    Friend WithEvents dgvPreviewExcel As DataGridView
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
    Friend WithEvents grbReportOption As GroupBox
    Friend WithEvents chkProductExpirationReport As CheckBox
    Friend WithEvents chkExtendedReportZ As CheckBox
    Friend WithEvents chkClearOrderPoint As CheckBox
    Friend WithEvents chkCombineReport As CheckBox
    Friend WithEvents chkPOImport As CheckBox
    Friend WithEvents chkMultiReport As CheckBox
    Friend WithEvents grbPrintSetup As GroupBox
    Friend WithEvents grbWorkSheet As GroupBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents cboPaperSize As ComboBox
    Friend WithEvents cboPaperOrientation As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtZoom As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents txtPrintAreaColumn2 As TextBox
    Friend WithEvents txtPrintAreaColumn1 As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents txtSortLevel3 As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents txtSortLevel2 As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents txtSortLevel1 As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents txtRangeEnd As TextBox
    Friend WithEvents txtRangeStart As TextBox
    Friend WithEvents lsvName As ColumnHeader
    Friend WithEvents lsvDateModified As ColumnHeader
    Friend WithEvents lsvFilePath As ColumnHeader
    Friend WithEvents ToolStripStatusLabel2 As ToolStripStatusLabel
    Friend WithEvents lblRows As ToolStripStatusLabel
    Friend WithEvents lblColumns As ToolStripStatusLabel
    Friend WithEvents lblProcessing As ToolStripStatusLabel
    Friend WithEvents tpbExport As ToolStripProgressBar
    Friend WithEvents lblCombineReport As ToolStripStatusLabel
End Class
