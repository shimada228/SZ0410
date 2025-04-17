<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class SZ0415FRM
#Region "Windows フォーム デザイナによって生成されたコード "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'この呼び出しは、Windows フォーム デザイナで必要です。
		InitializeComponent()
	End Sub
	'Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows フォーム デザイナで必要です。
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents SPRD As AxFPSpreadADO.AxfpSpread
	Public WithEvents CMB030 As System.Windows.Forms.ComboBox
	Public WithEvents CMB020 As System.Windows.Forms.ComboBox
	Public WithEvents CMB010 As System.Windows.Forms.ComboBox
	Public WithEvents _CMDOFNC_12 As Control.AximButton6.AximButton
	Public WithEvents _CMDOFNC_0 As Control.AximButton6.AximButton
	Public WithEvents _CMDOFNC_5 As Control.AximButton6.AximButton
	Public WithEvents CMDODSP As Control.AximButton6.AximButton
	Public WithEvents LBL040T As System.Windows.Forms.Label
	Public WithEvents LBL030T As System.Windows.Forms.Label
	Public WithEvents LBL020T As System.Windows.Forms.Label
	Public WithEvents _LBLFNC_5 As System.Windows.Forms.Label
	Public WithEvents _LBLFNC_12 As System.Windows.Forms.Label
	Public WithEvents _LBLFNC_0 As System.Windows.Forms.Label
	Public WithEvents LBL010T As System.Windows.Forms.Label
	Public WithEvents CMDOFNC As Control.AximButtonArray.AximButton
	Public WithEvents LBLFNC As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(SZ0415FRM))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.SPRD = New AxFPSpreadADO.AxfpSpread
		Me.CMB030 = New System.Windows.Forms.ComboBox
		Me.CMB020 = New System.Windows.Forms.ComboBox
		Me.CMB010 = New System.Windows.Forms.ComboBox
		Me._CMDOFNC_12 = New Control.AximButton6.AximButton
		Me._CMDOFNC_0 = New Control.AximButton6.AximButton
		Me._CMDOFNC_5 = New Control.AximButton6.AximButton
		Me.CMDODSP = New Control.AximButton6.AximButton
		Me.LBL040T = New System.Windows.Forms.Label
		Me.LBL030T = New System.Windows.Forms.Label
		Me.LBL020T = New System.Windows.Forms.Label
		Me._LBLFNC_5 = New System.Windows.Forms.Label
		Me._LBLFNC_12 = New System.Windows.Forms.Label
		Me._LBLFNC_0 = New System.Windows.Forms.Label
		Me.LBL010T = New System.Windows.Forms.Label
		Me.CMDOFNC = New Control.AximButtonArray.AximButton(components)
		Me.LBLFNC = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.SPRD, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._CMDOFNC_12, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._CMDOFNC_0, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._CMDOFNC_5, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.CMDODSP, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.CMDOFNC, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.LBLFNC, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "JAN商品分類検索"
		Me.ClientSize = New System.Drawing.Size(442, 430)
		Me.Location = New System.Drawing.Point(189, 153)
		Me.Font = New System.Drawing.Font("ＭＳ 明朝", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me.Icon = CType(resources.GetObject("SZ0415FRM.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "SZ0415FRM"
		SPRD.OcxState = CType(resources.GetObject("SPRD.OcxState"), System.Windows.Forms.AxHost.State)
		Me.SPRD.Size = New System.Drawing.Size(368, 262)
		Me.SPRD.Location = New System.Drawing.Point(60, 118)
		Me.SPRD.TabIndex = 4
		Me.SPRD.Name = "SPRD"
		Me.CMB030.BackColor = System.Drawing.Color.White
		Me.CMB030.Size = New System.Drawing.Size(367, 21)
		Me.CMB030.Location = New System.Drawing.Point(60, 52)
		Me.CMB030.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.CMB030.TabIndex = 2
		Me.CMB030.CausesValidation = True
		Me.CMB030.Enabled = True
		Me.CMB030.ForeColor = System.Drawing.SystemColors.WindowText
		Me.CMB030.IntegralHeight = True
		Me.CMB030.Cursor = System.Windows.Forms.Cursors.Default
		Me.CMB030.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CMB030.Sorted = False
		Me.CMB030.TabStop = True
		Me.CMB030.Visible = True
		Me.CMB030.Name = "CMB030"
		Me.CMB020.BackColor = System.Drawing.Color.White
		Me.CMB020.Size = New System.Drawing.Size(367, 21)
		Me.CMB020.Location = New System.Drawing.Point(60, 30)
		Me.CMB020.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.CMB020.TabIndex = 1
		Me.CMB020.CausesValidation = True
		Me.CMB020.Enabled = True
		Me.CMB020.ForeColor = System.Drawing.SystemColors.WindowText
		Me.CMB020.IntegralHeight = True
		Me.CMB020.Cursor = System.Windows.Forms.Cursors.Default
		Me.CMB020.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CMB020.Sorted = False
		Me.CMB020.TabStop = True
		Me.CMB020.Visible = True
		Me.CMB020.Name = "CMB020"
		Me.CMB010.BackColor = System.Drawing.Color.White
		Me.CMB010.Size = New System.Drawing.Size(367, 21)
		Me.CMB010.Location = New System.Drawing.Point(60, 8)
		Me.CMB010.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.CMB010.TabIndex = 0
		Me.CMB010.CausesValidation = True
		Me.CMB010.Enabled = True
		Me.CMB010.ForeColor = System.Drawing.SystemColors.WindowText
		Me.CMB010.IntegralHeight = True
		Me.CMB010.Cursor = System.Windows.Forms.Cursors.Default
		Me.CMB010.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CMB010.Sorted = False
		Me.CMB010.TabStop = True
		Me.CMB010.Visible = True
		Me.CMB010.Name = "CMB010"
		_CMDOFNC_12.OcxState = CType(resources.GetObject("_CMDOFNC_12.OcxState"), System.Windows.Forms.AxHost.State)
		Me._CMDOFNC_12.Size = New System.Drawing.Size(69, 21)
		Me._CMDOFNC_12.Location = New System.Drawing.Point(359, 403)
		Me._CMDOFNC_12.TabIndex = 5
		Me._CMDOFNC_12.Name = "_CMDOFNC_12"
		_CMDOFNC_0.OcxState = CType(resources.GetObject("_CMDOFNC_0.OcxState"), System.Windows.Forms.AxHost.State)
		Me._CMDOFNC_0.Size = New System.Drawing.Size(65, 21)
		Me._CMDOFNC_0.Location = New System.Drawing.Point(5, 403)
		Me._CMDOFNC_0.TabIndex = 7
		Me._CMDOFNC_0.Name = "_CMDOFNC_0"
		_CMDOFNC_5.OcxState = CType(resources.GetObject("_CMDOFNC_5.OcxState"), System.Windows.Forms.AxHost.State)
		Me._CMDOFNC_5.Size = New System.Drawing.Size(65, 21)
		Me._CMDOFNC_5.Location = New System.Drawing.Point(71, 403)
		Me._CMDOFNC_5.TabIndex = 6
		Me._CMDOFNC_5.Name = "_CMDOFNC_5"
		CMDODSP.OcxState = CType(resources.GetObject("CMDODSP.OcxState"), System.Windows.Forms.AxHost.State)
		Me.CMDODSP.Size = New System.Drawing.Size(69, 21)
		Me.CMDODSP.Location = New System.Drawing.Point(358, 80)
		Me.CMDODSP.TabIndex = 3
		Me.CMDODSP.Name = "CMDODSP"
		Me.LBL040T.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.LBL040T.Text = "細分類："
		Me.LBL040T.Font = New System.Drawing.Font("ＭＳ 明朝", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me.LBL040T.Size = New System.Drawing.Size(53, 13)
		Me.LBL040T.Location = New System.Drawing.Point(8, 120)
		Me.LBL040T.TabIndex = 14
		Me.LBL040T.BackColor = System.Drawing.SystemColors.Control
		Me.LBL040T.Enabled = True
		Me.LBL040T.ForeColor = System.Drawing.SystemColors.ControlText
		Me.LBL040T.Cursor = System.Windows.Forms.Cursors.Default
		Me.LBL040T.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.LBL040T.UseMnemonic = True
		Me.LBL040T.Visible = True
		Me.LBL040T.AutoSize = False
		Me.LBL040T.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.LBL040T.Name = "LBL040T"
		Me.LBL030T.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.LBL030T.Text = "小分類："
		Me.LBL030T.Font = New System.Drawing.Font("ＭＳ 明朝", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me.LBL030T.Size = New System.Drawing.Size(53, 13)
		Me.LBL030T.Location = New System.Drawing.Point(8, 56)
		Me.LBL030T.TabIndex = 13
		Me.LBL030T.BackColor = System.Drawing.SystemColors.Control
		Me.LBL030T.Enabled = True
		Me.LBL030T.ForeColor = System.Drawing.SystemColors.ControlText
		Me.LBL030T.Cursor = System.Windows.Forms.Cursors.Default
		Me.LBL030T.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.LBL030T.UseMnemonic = True
		Me.LBL030T.Visible = True
		Me.LBL030T.AutoSize = False
		Me.LBL030T.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.LBL030T.Name = "LBL030T"
		Me.LBL020T.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.LBL020T.Text = "中分類："
		Me.LBL020T.Font = New System.Drawing.Font("ＭＳ 明朝", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me.LBL020T.Size = New System.Drawing.Size(53, 13)
		Me.LBL020T.Location = New System.Drawing.Point(8, 34)
		Me.LBL020T.TabIndex = 12
		Me.LBL020T.BackColor = System.Drawing.SystemColors.Control
		Me.LBL020T.Enabled = True
		Me.LBL020T.ForeColor = System.Drawing.SystemColors.ControlText
		Me.LBL020T.Cursor = System.Windows.Forms.Cursors.Default
		Me.LBL020T.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.LBL020T.UseMnemonic = True
		Me.LBL020T.Visible = True
		Me.LBL020T.AutoSize = False
		Me.LBL020T.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.LBL020T.Name = "LBL020T"
		Me._LBLFNC_5.Text = "F5"
		Me._LBLFNC_5.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me._LBLFNC_5.Size = New System.Drawing.Size(21, 13)
		Me._LBLFNC_5.Location = New System.Drawing.Point(72, 390)
		Me._LBLFNC_5.TabIndex = 11
		Me._LBLFNC_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._LBLFNC_5.BackColor = System.Drawing.SystemColors.Control
		Me._LBLFNC_5.Enabled = True
		Me._LBLFNC_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._LBLFNC_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._LBLFNC_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._LBLFNC_5.UseMnemonic = True
		Me._LBLFNC_5.Visible = True
		Me._LBLFNC_5.AutoSize = False
		Me._LBLFNC_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._LBLFNC_5.Name = "_LBLFNC_5"
		Me._LBLFNC_12.Text = "F12"
		Me._LBLFNC_12.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me._LBLFNC_12.Size = New System.Drawing.Size(21, 13)
		Me._LBLFNC_12.Location = New System.Drawing.Point(360, 390)
		Me._LBLFNC_12.TabIndex = 10
		Me._LBLFNC_12.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._LBLFNC_12.BackColor = System.Drawing.SystemColors.Control
		Me._LBLFNC_12.Enabled = True
		Me._LBLFNC_12.ForeColor = System.Drawing.SystemColors.ControlText
		Me._LBLFNC_12.Cursor = System.Windows.Forms.Cursors.Default
		Me._LBLFNC_12.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._LBLFNC_12.UseMnemonic = True
		Me._LBLFNC_12.Visible = True
		Me._LBLFNC_12.AutoSize = False
		Me._LBLFNC_12.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._LBLFNC_12.Name = "_LBLFNC_12"
		Me._LBLFNC_0.Text = "ESC"
		Me._LBLFNC_0.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me._LBLFNC_0.Size = New System.Drawing.Size(21, 13)
		Me._LBLFNC_0.Location = New System.Drawing.Point(6, 390)
		Me._LBLFNC_0.TabIndex = 8
		Me._LBLFNC_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._LBLFNC_0.BackColor = System.Drawing.SystemColors.Control
		Me._LBLFNC_0.Enabled = True
		Me._LBLFNC_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._LBLFNC_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._LBLFNC_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._LBLFNC_0.UseMnemonic = True
		Me._LBLFNC_0.Visible = True
		Me._LBLFNC_0.AutoSize = False
		Me._LBLFNC_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._LBLFNC_0.Name = "_LBLFNC_0"
		Me.LBL010T.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.LBL010T.Text = "大分類："
		Me.LBL010T.Font = New System.Drawing.Font("ＭＳ 明朝", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me.LBL010T.Size = New System.Drawing.Size(53, 13)
		Me.LBL010T.Location = New System.Drawing.Point(8, 12)
		Me.LBL010T.TabIndex = 9
		Me.LBL010T.BackColor = System.Drawing.SystemColors.Control
		Me.LBL010T.Enabled = True
		Me.LBL010T.ForeColor = System.Drawing.SystemColors.ControlText
		Me.LBL010T.Cursor = System.Windows.Forms.Cursors.Default
		Me.LBL010T.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.LBL010T.UseMnemonic = True
		Me.LBL010T.Visible = True
		Me.LBL010T.AutoSize = False
		Me.LBL010T.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.LBL010T.Name = "LBL010T"
		Me.Controls.Add(SPRD)
		Me.Controls.Add(CMB030)
		Me.Controls.Add(CMB020)
		Me.Controls.Add(CMB010)
		Me.Controls.Add(_CMDOFNC_12)
		Me.Controls.Add(_CMDOFNC_0)
		Me.Controls.Add(_CMDOFNC_5)
		Me.Controls.Add(CMDODSP)
		Me.Controls.Add(LBL040T)
		Me.Controls.Add(LBL030T)
		Me.Controls.Add(LBL020T)
		Me.Controls.Add(_LBLFNC_5)
		Me.Controls.Add(_LBLFNC_12)
		Me.Controls.Add(_LBLFNC_0)
		Me.Controls.Add(LBL010T)
		Me.CMDOFNC.SetIndex(_CMDOFNC_12, CType(12, Short))
		Me.CMDOFNC.SetIndex(_CMDOFNC_0, CType(0, Short))
		Me.CMDOFNC.SetIndex(_CMDOFNC_5, CType(5, Short))
		Me.LBLFNC.SetIndex(_LBLFNC_5, CType(5, Short))
		Me.LBLFNC.SetIndex(_LBLFNC_12, CType(12, Short))
		Me.LBLFNC.SetIndex(_LBLFNC_0, CType(0, Short))
		CType(Me.LBLFNC, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.CMDOFNC, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.CMDODSP, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._CMDOFNC_5, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._CMDOFNC_0, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._CMDOFNC_12, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SPRD, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class