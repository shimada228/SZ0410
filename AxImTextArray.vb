'UPGRADE_WARNING: ActiveX コントロール配列を含むフォームを表示するには、プロジェクト全体をコンパイルする必要があります。

Imports System.ComponentModel

<ProvideProperty("Index",GetType(AxOsktxtLibV5.AxImText))> Public Class AxImTextArray
	Inherits Microsoft.VisualBasic.Compatibility.VB6.BaseOcxArray
	Implements IExtenderProvider

	Public Sub New()
		MyBase.New()
	End Sub

	Public Sub New(ByVal Container As IContainer)
		MyBase.New(Container)
	End Sub

	Public Shadows Event [Change] (ByVal sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event [ClickEvent] (ByVal sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event [DblClick] (ByVal sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event [InputPastEnd] (ByVal sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event [InvalidKey] (ByVal sender As System.Object, ByVal e As AxOsktxtLibV5.__ImText_InvalidKeyEvent)
	Public Shadows Event [KeyDownEvent] (ByVal sender As System.Object, ByVal e As AxOsktxtLibV5.__ImText_KeyDownEvent)
	Public Shadows Event [KeyPressEvent] (ByVal sender As System.Object, ByVal e As AxOsktxtLibV5.__ImText_KeyPressEvent)
	Public Shadows Event [KeyUpEvent] (ByVal sender As System.Object, ByVal e As AxOsktxtLibV5.__ImText_KeyUpEvent)
	Public Shadows Event [MouseDownEvent] (ByVal sender As System.Object, ByVal e As AxOsktxtLibV5.__ImText_MouseDownEvent)
	Public Shadows Event [MouseMoveEvent] (ByVal sender As System.Object, ByVal e As AxOsktxtLibV5.__ImText_MouseMoveEvent)
	Public Shadows Event [MouseUpEvent] (ByVal sender As System.Object, ByVal e As AxOsktxtLibV5.__ImText_MouseUpEvent)
	Public Shadows Event [OutOfRange] (ByVal sender As System.Object, ByVal e As System.EventArgs)

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Function CanExtend(ByVal target As Object) As Boolean Implements IExtenderProvider.CanExtend
		If TypeOf target Is AxOsktxtLibV5.AxImText Then
			Return BaseCanExtend(target)
		End If
	End Function

	Public Function GetIndex(ByVal o As AxOsktxtLibV5.AxImText) As Short
		Return BaseGetIndex(o)
	End Function

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Sub SetIndex(ByVal o As AxOsktxtLibV5.AxImText, ByVal Index As Short)
		BaseSetIndex(o, Index)
	End Sub

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Function ShouldSerializeIndex(ByVal o As AxOsktxtLibV5.AxImText) As Boolean
		Return BaseShouldSerializeIndex(o)
	End Function

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Sub ResetIndex(ByVal o As AxOsktxtLibV5.AxImText)
		BaseResetIndex(o)
	End Sub

	Default Public ReadOnly Property Item(ByVal Index As Short) As AxOsktxtLibV5.AxImText
		Get
			Item = CType(BaseGetItem(Index), AxOsktxtLibV5.AxImText)
		End Get
	End Property

	Protected Overrides Function GetControlInstanceType() As System.Type
		Return GetType(AxOsktxtLibV5.AxImText)
	End Function

	Protected Overrides Sub HookUpControlEvents(ByVal o As Object)
		Dim ctl As AxOsktxtLibV5.AxImText = CType(o, AxOsktxtLibV5.AxImText)
		MyBase.HookUpControlEvents(o)
		If Not ChangeEvent Is Nothing Then
			AddHandler ctl.Change, New System.EventHandler(AddressOf HandleChange)
		End If
		If Not ClickEventEvent Is Nothing Then
			AddHandler ctl.ClickEvent, New System.EventHandler(AddressOf HandleClickEvent)
		End If
		If Not DblClickEvent Is Nothing Then
			AddHandler ctl.DblClick, New System.EventHandler(AddressOf HandleDblClick)
		End If
		If Not InputPastEndEvent Is Nothing Then
			AddHandler ctl.InputPastEnd, New System.EventHandler(AddressOf HandleInputPastEnd)
		End If
		If Not InvalidKeyEvent Is Nothing Then
			AddHandler ctl.InvalidKey, New AxOsktxtLibV5.__ImText_InvalidKeyEventHandler(AddressOf HandleInvalidKey)
		End If
		If Not KeyDownEventEvent Is Nothing Then
			AddHandler ctl.KeyDownEvent, New AxOsktxtLibV5.__ImText_KeyDownEventHandler(AddressOf HandleKeyDownEvent)
		End If
		If Not KeyPressEventEvent Is Nothing Then
			AddHandler ctl.KeyPressEvent, New AxOsktxtLibV5.__ImText_KeyPressEventHandler(AddressOf HandleKeyPressEvent)
		End If
		If Not KeyUpEventEvent Is Nothing Then
			AddHandler ctl.KeyUpEvent, New AxOsktxtLibV5.__ImText_KeyUpEventHandler(AddressOf HandleKeyUpEvent)
		End If
		If Not MouseDownEventEvent Is Nothing Then
			AddHandler ctl.MouseDownEvent, New AxOsktxtLibV5.__ImText_MouseDownEventHandler(AddressOf HandleMouseDownEvent)
		End If
		If Not MouseMoveEventEvent Is Nothing Then
			AddHandler ctl.MouseMoveEvent, New AxOsktxtLibV5.__ImText_MouseMoveEventHandler(AddressOf HandleMouseMoveEvent)
		End If
		If Not MouseUpEventEvent Is Nothing Then
			AddHandler ctl.MouseUpEvent, New AxOsktxtLibV5.__ImText_MouseUpEventHandler(AddressOf HandleMouseUpEvent)
		End If
		If Not OutOfRangeEvent Is Nothing Then
			AddHandler ctl.OutOfRange, New System.EventHandler(AddressOf HandleOutOfRange)
		End If
	End Sub

	Private Sub HandleChange (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [Change] (sender, e)
	End Sub

	Private Sub HandleClickEvent (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [ClickEvent] (sender, e)
	End Sub

	Private Sub HandleDblClick (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [DblClick] (sender, e)
	End Sub

	Private Sub HandleInputPastEnd (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [InputPastEnd] (sender, e)
	End Sub

	Private Sub HandleInvalidKey (ByVal sender As System.Object, ByVal e As AxOsktxtLibV5.__ImText_InvalidKeyEvent) 
		RaiseEvent [InvalidKey] (sender, e)
	End Sub

	Private Sub HandleKeyDownEvent (ByVal sender As System.Object, ByVal e As AxOsktxtLibV5.__ImText_KeyDownEvent) 
		RaiseEvent [KeyDownEvent] (sender, e)
	End Sub

	Private Sub HandleKeyPressEvent (ByVal sender As System.Object, ByVal e As AxOsktxtLibV5.__ImText_KeyPressEvent) 
		RaiseEvent [KeyPressEvent] (sender, e)
	End Sub

	Private Sub HandleKeyUpEvent (ByVal sender As System.Object, ByVal e As AxOsktxtLibV5.__ImText_KeyUpEvent) 
		RaiseEvent [KeyUpEvent] (sender, e)
	End Sub

	Private Sub HandleMouseDownEvent (ByVal sender As System.Object, ByVal e As AxOsktxtLibV5.__ImText_MouseDownEvent) 
		RaiseEvent [MouseDownEvent] (sender, e)
	End Sub

	Private Sub HandleMouseMoveEvent (ByVal sender As System.Object, ByVal e As AxOsktxtLibV5.__ImText_MouseMoveEvent) 
		RaiseEvent [MouseMoveEvent] (sender, e)
	End Sub

	Private Sub HandleMouseUpEvent (ByVal sender As System.Object, ByVal e As AxOsktxtLibV5.__ImText_MouseUpEvent) 
		RaiseEvent [MouseUpEvent] (sender, e)
	End Sub

	Private Sub HandleOutOfRange (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [OutOfRange] (sender, e)
	End Sub

End Class

