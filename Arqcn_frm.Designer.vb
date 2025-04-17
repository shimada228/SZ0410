<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class ARQCNFRM
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
	Public WithEvents TXT030 As System.Windows.Forms.TextBox
	Public WithEvents TXT010 As System.Windows.Forms.TextBox
	Public WithEvents TXT020 As System.Windows.Forms.TextBox
	Public WithEvents CMDO020 As Control.AximButton6.AximButton
	Public WithEvents CMDO010 As Control.AximButton6.AximButton
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(ARQCNFRM))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.TXT030 = New System.Windows.Forms.TextBox
		Me.TXT010 = New System.Windows.Forms.TextBox
		Me.TXT020 = New System.Windows.Forms.TextBox
		Me.CMDO020 = New Control.AximButton6.AximButton
		Me.CMDO010 = New Control.AximButton6.AximButton
		Me._Label1_2 = New System.Windows.Forms.Label
		Me._Label1_1 = New System.Windows.Forms.Label
		Me._Label1_0 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.CMDO020, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.CMDO010, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "データベースに接続します"
		Me.ClientSize = New System.Drawing.Size(340, 143)
		Me.Location = New System.Drawing.Point(217, 311)
		Me.Font = New System.Drawing.Font("ＭＳ 明朝", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me.ForeColor = System.Drawing.SystemColors.WindowText
		Me.KeyPreview = True
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "ARQCNFRM"
		Me.TXT030.AutoSize = False
		Me.TXT030.Font = New System.Drawing.Font("ＭＳ 明朝", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me.TXT030.Size = New System.Drawing.Size(178, 22)
		Me.TXT030.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.TXT030.Location = New System.Drawing.Point(132, 56)
		Me.TXT030.Maxlength = 30
		Me.TXT030.TabIndex = 7
		Me.TXT030.AcceptsReturn = True
		Me.TXT030.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.TXT030.BackColor = System.Drawing.SystemColors.Window
		Me.TXT030.CausesValidation = True
		Me.TXT030.Enabled = True
		Me.TXT030.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TXT030.HideSelection = True
		Me.TXT030.ReadOnly = False
		Me.TXT030.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TXT030.MultiLine = False
		Me.TXT030.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TXT030.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TXT030.TabStop = True
		Me.TXT030.Visible = True
		Me.TXT030.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TXT030.Name = "TXT030"
		Me.TXT010.AutoSize = False
		Me.TXT010.Font = New System.Drawing.Font("ＭＳ 明朝", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me.TXT010.Size = New System.Drawing.Size(178, 22)
		Me.TXT010.IMEMode = System.Windows.Forms.ImeMode.Off
		Me.TXT010.Location = New System.Drawing.Point(132, 6)
		Me.TXT010.Maxlength = 30
		Me.TXT010.TabIndex = 0
		Me.TXT010.AcceptsReturn = True
		Me.TXT010.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.TXT010.BackColor = System.Drawing.SystemColors.Window
		Me.TXT010.CausesValidation = True
		Me.TXT010.Enabled = True
		Me.TXT010.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TXT010.HideSelection = True
		Me.TXT010.ReadOnly = False
		Me.TXT010.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TXT010.MultiLine = False
		Me.TXT010.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TXT010.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TXT010.TabStop = True
		Me.TXT010.Visible = True
		Me.TXT010.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TXT010.Name = "TXT010"
		Me.TXT020.AutoSize = False
		Me.TXT020.Font = New System.Drawing.Font("ＭＳ 明朝", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me.TXT020.Size = New System.Drawing.Size(178, 22)
		Me.TXT020.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.TXT020.Location = New System.Drawing.Point(132, 32)
		Me.TXT020.Maxlength = 30
		Me.TXT020.PasswordChar = ChrW(42)
		Me.TXT020.TabIndex = 1
		Me.TXT020.AcceptsReturn = True
		Me.TXT020.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.TXT020.BackColor = System.Drawing.SystemColors.Window
		Me.TXT020.CausesValidation = True
		Me.TXT020.Enabled = True
		Me.TXT020.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TXT020.HideSelection = True
		Me.TXT020.ReadOnly = False
		Me.TXT020.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TXT020.MultiLine = False
		Me.TXT020.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TXT020.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TXT020.TabStop = True
		Me.TXT020.Visible = True
		Me.TXT020.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TXT020.Name = "TXT020"
		Me.CMDO020.Size = New System.Drawing.Size(94, 26)
		Me.CMDO020.Location = New System.Drawing.Point(196, 102)
		Me.CMDO020.TabIndex = 3
		Me.CMDO020.Name = "CMDO020"
		Me.CMDO010.Size = New System.Drawing.Size(94, 26)
		Me.CMDO010.Location = New System.Drawing.Point(60, 102)
		Me.CMDO010.TabIndex = 2
		Me.CMDO010.Name = "CMDO010"
		Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_2.Text = "データベース名:"
		Me._Label1_2.Size = New System.Drawing.Size(114, 17)
		Me._Label1_2.Location = New System.Drawing.Point(12, 62)
		Me._Label1_2.TabIndex = 6
		Me._Label1_2.BackColor = System.Drawing.Color.Transparent
		Me._Label1_2.Enabled = True
		Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_2.UseMnemonic = True
		Me._Label1_2.Visible = True
		Me._Label1_2.AutoSize = False
		Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_2.Name = "_Label1_2"
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_1.BackColor = System.Drawing.Color.Transparent
		Me._Label1_1.Text = "パスワード:"
		Me._Label1_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Label1_1.Size = New System.Drawing.Size(114, 17)
		Me._Label1_1.Location = New System.Drawing.Point(12, 39)
		Me._Label1_1.TabIndex = 5
		Me._Label1_1.Enabled = True
		Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_1.UseMnemonic = True
		Me._Label1_1.Visible = True
		Me._Label1_1.AutoSize = False
		Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_1.Name = "_Label1_1"
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_0.BackColor = System.Drawing.Color.Transparent
		Me._Label1_0.Text = "ユーザ名:"
		Me._Label1_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Label1_0.Size = New System.Drawing.Size(114, 17)
		Me._Label1_0.Location = New System.Drawing.Point(12, 12)
		Me._Label1_0.TabIndex = 4
		Me._Label1_0.Enabled = True
		Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_0.UseMnemonic = True
		Me._Label1_0.Visible = True
		Me._Label1_0.AutoSize = False
		Me._Label1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_0.Name = "_Label1_0"
		Me.Controls.Add(TXT030)
		Me.Controls.Add(TXT010)
		Me.Controls.Add(TXT020)
		Me.Controls.Add(CMDO020)
		Me.Controls.Add(CMDO010)
		Me.Controls.Add(_Label1_2)
		Me.Controls.Add(_Label1_1)
		Me.Controls.Add(_Label1_0)
		Me.Label1.SetIndex(_Label1_2, CType(2, Short))
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.CMDO010, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.CMDO020, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class