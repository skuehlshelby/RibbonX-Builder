Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Builders

	Public NotInheritable Class DropDownBuilder
		Implements IBuilder(Of DropDown)
		Implements IID(Of DropDownBuilder)
		Implements IAddButton(Of DropDownBuilder)
		Implements IEnabled(Of DropDownBuilder)
		Implements IInsert(Of DropDownBuilder)
		Implements IVisible(Of DropDownBuilder)
		Implements ILabelScreenTipSuperTip(Of DropDownBuilder)
		Implements IKeyTip(Of DropDownBuilder)
		Implements IShowLabel(Of DropDownBuilder)
		Implements IImage(Of DropDownBuilder)
		Implements IShowImage(Of DropDownBuilder)
		Implements ISizeString(Of DropDownBuilder)
		Implements IGetItemId(Of DropDownBuilder)
		Implements IGetItemCount(Of DropDownBuilder)
		Implements IGetItemLabel(Of DropDownBuilder)
		Implements IShowItemLabel(Of DropDownBuilder)
		Implements IGetItemScreentip(Of DropDownBuilder)
		Implements IGetItemSupertip(Of DropDownBuilder)
		Implements IGetItemImage(Of DropDownBuilder)
		Implements IShowItemImage(Of DropDownBuilder)
		Implements IGetSelectedItemId(Of DropDownBuilder)
		Implements IGetSelectedItemIndex(Of DropDownBuilder)

		Private ReadOnly _builder As ControlBuilder
		Private ReadOnly _buttons As ICollection(Of Button) = New List(Of Button)
		Private ReadOnly _beforeSelectionChangeHandlers As ICollection(Of EventHandler(Of DropDown.BeforeSelectionChangeEventArgs)) = New List(Of EventHandler(Of DropDown.BeforeSelectionChangeEventArgs))
		Private ReadOnly _selectionChangedHandlers As ICollection(Of EventHandler(Of DropDown.SelectionChangedEventArgs)) = New List(Of EventHandler(Of DropDown.SelectionChangedEventArgs))

		Public Sub New()
			Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of DropDown)
			Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
			attributeGroupBuilder.SetDefaults(defaultProvider)
			_builder = New ControlBuilder(attributeGroupBuilder)
		End Sub

		Public Sub New(template As DropDown)
			Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
			attributeGroupBuilder.SetDefaults(template)
			attributeGroupBuilder.AddID(IdManager.GetID(Of DropDown)())
			_builder = New ControlBuilder(attributeGroupBuilder)

			Array.ForEach(template.GetBeforeSelectionChangeInvocationList(), Sub(handler) _beforeSelectionChangeHandlers.Add(handler))
			Array.ForEach(template.GetSelectionChangedInvocationList(), Sub(handler) _selectionChangedHandlers.Add(handler))
		End Sub

		Public Sub New(template As RibbonElement)
			Dim defaultProvider As IDefaultProvider = TryCast(template, IDefaultProvider)

			If defaultProvider IsNot Nothing Then
				Dim templateAttributes As AttributeSet = defaultProvider.GetDefaults()
				Dim dropDownAttributes As AttributeSet = New DefaultProvider(Of DropDown)().GetDefaults()
				Dim intersection As AttributeSet = New AttributeSet(templateAttributes.Where(Function(a) dropDownAttributes.Contains(a)))
				Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder(intersection)
				attributeGroupBuilder.AddID(IdManager.GetID(Of DropDown)())
				_builder = New ControlBuilder(attributeGroupBuilder)
			Else
				Throw New ArgumentException($"Could not copy applicable properties of type '{template.GetType().Name}' to type '{GetType(DropDown)}'")
			End If
		End Sub

		Public Function Add(button As Button) As DropDownBuilder Implements IAddButton(Of DropDownBuilder).Add
			_buttons.Add(button)
			Return Me
		End Function

		Public Function AfterSelectionChanges(action As Action(Of DropDown.SelectionChangedEventArgs)) As DropDownBuilder
			Dim handler As SelectionChangeEventHandlerHelper = New SelectionChangeEventHandlerHelper(action)

			_selectionChangedHandlers.Add(AddressOf handler.Handle)

			Return Me
		End Function

		Public Function AsWideAs(sizeString As String) As DropDownBuilder Implements ISizeString(Of DropDownBuilder).AsWideAs
			_builder.WithSize(sizeString)
			Return Me
		End Function

		Public Function BeforeSelectionChanges(action As Action(Of DropDown.BeforeSelectionChangeEventArgs)) As DropDownBuilder
			Dim handler As BeforeSelectionChangeEventHandlerHelper = New BeforeSelectionChangeEventHandlerHelper(action)

			_beforeSelectionChangeHandlers.Add(AddressOf handler.Handle)

			Return Me
		End Function

		Public Function Build(Optional tag As Object = Nothing) As DropDown Implements IBuilder(Of DropDown).Build
			Dim dropDown As DropDown = New DropDown(_builder.Build(), _buttons.ToArray(), tag:=tag)

			For Each handler As EventHandler(Of DropDown.BeforeSelectionChangeEventArgs) In _beforeSelectionChangeHandlers
				AddHandler dropDown.BeforeSelectionChange, handler
			Next

			For Each handler As EventHandler(Of DropDown.SelectionChangedEventArgs) In _selectionChangedHandlers
				AddHandler dropDown.SelectionChanged, handler
			Next

			Return dropDown
		End Function

		Public Function Disabled() As DropDownBuilder Implements IEnabled(Of DropDownBuilder).Disabled
			_builder.Disabled()
			Return Me
		End Function

		Public Function Disabled(callback As FromControl(Of Boolean)) As DropDownBuilder Implements IEnabled(Of DropDownBuilder).Disabled
			_builder.Disabled(callback)
			Return Me
		End Function

		Public Function Enabled() As DropDownBuilder Implements IEnabled(Of DropDownBuilder).Enabled
			_builder.Enabled()
			Return Me
		End Function

		Public Function Enabled(callback As FromControl(Of Boolean)) As DropDownBuilder Implements IEnabled(Of DropDownBuilder).Enabled
			_builder.Enabled(callback)
			Return Me
		End Function

		Public Function GetItemCountFrom(callback As FromControl(Of Integer)) As DropDownBuilder Implements IGetItemCount(Of DropDownBuilder).GetItemCountFrom
			_builder.GetItemCountFrom(callback)
			Return Me
		End Function

		Public Function GetItemIdFrom(callback As FromControlAndIndex(Of String)) As DropDownBuilder Implements IGetItemId(Of DropDownBuilder).GetItemIdFrom
			_builder.GetItemIDFrom(callback)
			Return Me
		End Function

		Public Function GetItemImageFrom(callback As FromControlAndIndex(Of IPictureDisp)) As DropDownBuilder Implements IGetItemImage(Of DropDownBuilder).GetItemImageFrom
			_builder.GetItemImageFrom(callback)
			Return Me
		End Function

		Public Function GetItemLabelFrom(callback As FromControlAndIndex(Of String)) As DropDownBuilder Implements IGetItemLabel(Of DropDownBuilder).GetItemLabelFrom
			_builder.GetItemLabelFrom(callback)
			Return Me
		End Function

		Public Function GetItemScreenTipFrom(callback As FromControlAndIndex(Of String)) As DropDownBuilder Implements IGetItemScreentip(Of DropDownBuilder).GetItemScreenTipFrom
			_builder.GetItemScreentipFrom(callback)
			Return Me
		End Function

		Public Function GetItemSuperTipFrom(callback As FromControlAndIndex(Of String)) As DropDownBuilder Implements IGetItemSupertip(Of DropDownBuilder).GetItemSuperTipFrom
			_builder.GetItemSupertipFrom(callback)
			Return Me
		End Function

		Public Function GetSelectedItemIdFrom(getSelected As FromControl(Of String), setSelected As SelectionChanged) As DropDownBuilder Implements IGetSelectedItemId(Of DropDownBuilder).GetSelectedItemIdFrom
			_builder.GetSelectedItemIDFrom(getSelected, setSelected)
			Return Me
		End Function

		Public Function GetSelectedItemIndexFrom(getSelected As FromControl(Of Integer), setSelected As SelectionChanged) As DropDownBuilder Implements IGetSelectedItemIndex(Of DropDownBuilder).GetSelectedItemIndexFrom
			_builder.GetSelectedItemIndexFrom(getSelected, setSelected)
			Return Me
		End Function

		Public Function HideImage() As DropDownBuilder Implements IShowImage(Of DropDownBuilder).HideImage
			_builder.HideImage()
			Return Me
		End Function

		Public Function HideImage(getShowImage As FromControl(Of Boolean)) As DropDownBuilder Implements IShowImage(Of DropDownBuilder).HideImage
			_builder.HideImage(getShowImage)
			Return Me
		End Function

		Public Function HideItemImages() As DropDownBuilder Implements IShowItemImage(Of DropDownBuilder).HideItemImages
			_builder.HideItemImages()
			Return Me
		End Function

		Public Function HideLabel() As DropDownBuilder Implements IShowLabel(Of DropDownBuilder).HideLabel
			_builder.HideLabel()
			Return Me
		End Function

		Public Function HideLabel(callback As FromControl(Of Boolean)) As DropDownBuilder Implements IShowLabel(Of DropDownBuilder).HideLabel
			_builder.HideLabel(callback)
			Return Me
		End Function

		Public Function HideLabelsOnChildItems() As DropDownBuilder Implements IShowItemLabel(Of DropDownBuilder).HideLabelsOnChildItems
			_builder.HideItemLabel()
			Return Me
		End Function

		Public Function InsertAfterMSO(builtInControl As MSO) As DropDownBuilder Implements IInsert(Of DropDownBuilder).InsertAfter
			_builder.InsertAfter(builtInControl)
			Return Me
		End Function

		Public Function InsertAfterQ(qualifiedControl As RibbonElement) As DropDownBuilder Implements IInsert(Of DropDownBuilder).InsertAfter
			_builder.InsertAfter(qualifiedControl)
			Return Me
		End Function

		Public Function InsertBeforeMSO(builtInControl As MSO) As DropDownBuilder Implements IInsert(Of DropDownBuilder).InsertBefore
			_builder.InsertBefore(builtInControl)
			Return Me
		End Function

		Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As DropDownBuilder Implements IInsert(Of DropDownBuilder).InsertBefore
			_builder.InsertBefore(qualifiedControl)
			Return Me
		End Function

		Public Function Invisible() As DropDownBuilder Implements IVisible(Of DropDownBuilder).Invisible
			_builder.Invisible()
			Return Me
		End Function

		Public Function Invisible(callback As FromControl(Of Boolean)) As DropDownBuilder Implements IVisible(Of DropDownBuilder).Invisible
			_builder.Invisible(callback)
			Return Me
		End Function

		Public Function ShowImage() As DropDownBuilder Implements IShowImage(Of DropDownBuilder).ShowImage
			_builder.ShowImage()
			Return Me
		End Function

		Public Function ShowImage(getShowImage As FromControl(Of Boolean)) As DropDownBuilder Implements IShowImage(Of DropDownBuilder).ShowImage
			_builder.ShowImage(getShowImage)
			Return Me
		End Function

		Public Function ShowItemImages() As DropDownBuilder Implements IShowItemImage(Of DropDownBuilder).ShowItemImages
			_builder.ShowItemImages()
			Return Me
		End Function

		Public Function ShowLabel() As DropDownBuilder Implements IShowLabel(Of DropDownBuilder).ShowLabel
			_builder.ShowLabel()
			Return Me
		End Function

		Public Function ShowLabel(callback As FromControl(Of Boolean)) As DropDownBuilder Implements IShowLabel(Of DropDownBuilder).ShowLabel
			_builder.ShowLabel(callback)
			Return Me
		End Function

		Public Function ShowLabelsOnChildItems() As DropDownBuilder Implements IShowItemLabel(Of DropDownBuilder).ShowLabelsOnChildItems
			_builder.ShowItemLabel()
			Return Me
		End Function

		Public Function Visible() As DropDownBuilder Implements IVisible(Of DropDownBuilder).Visible
			_builder.Visible()
			Return Me
		End Function

		Public Function Visible(callback As FromControl(Of Boolean)) As DropDownBuilder Implements IVisible(Of DropDownBuilder).Visible
			_builder.Visible(callback)
			Return Me
		End Function

		Public Function WithId(id As String) As DropDownBuilder Implements IID(Of DropDownBuilder).WithId
			_builder.WithId(id)
			Return Me
		End Function

		Public Function WithIdMso(mso As MSO) As DropDownBuilder Implements IID(Of DropDownBuilder).WithIdMso
			_builder.WithId(mso)
			Return Me
		End Function

		Public Function WithIdQ([namespace] As String, id As String) As DropDownBuilder Implements IID(Of DropDownBuilder).WithIdQ
			_builder.WithId([namespace], id)
			Return Me
		End Function

		Public Function WithImage(image As ImageMSO) As DropDownBuilder Implements IImage(Of DropDownBuilder).WithImage
			_builder.WithImage(image)
			Return Me
		End Function

		Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As DropDownBuilder Implements IImage(Of DropDownBuilder).WithImage
			_builder.WithImage(image, callback)
			Return Me
		End Function

		Public Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As DropDownBuilder Implements IImage(Of DropDownBuilder).WithImage
			_builder.WithImage(image, callback)
			Return Me
		End Function

		Public Function WithImage(imagePath As String) As DropDownBuilder Implements IImage(Of DropDownBuilder).WithImage
			_builder.WithImage(imagePath)
			Return Me
		End Function

		Public Function WithKeyTip(keyTip As KeyTip) As DropDownBuilder Implements IKeyTip(Of DropDownBuilder).WithKeyTip
			_builder.WithKeyTip(keyTip)
			Return Me
		End Function

		Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As DropDownBuilder Implements IKeyTip(Of DropDownBuilder).WithKeyTip
			_builder.WithKeyTip(keyTip, callback)
			Return Me
		End Function

		Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As DropDownBuilder Implements ILabelScreenTipSuperTip(Of DropDownBuilder).WithLabel
			_builder.WithLabel(label, copyToScreenTip)
			_builder.ShowLabel()
			Return Me
		End Function

		Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As DropDownBuilder Implements ILabelScreenTipSuperTip(Of DropDownBuilder).WithLabel
			_builder.WithLabel(label, callback, copyToScreenTip)
			_builder.ShowLabel()
			Return Me
		End Function

		Public Function WithScreenTip(screenTip As String) As DropDownBuilder Implements ILabelScreenTipSuperTip(Of DropDownBuilder).WithScreenTip
			_builder.WithScreentip(screenTip)
			Return Me
		End Function

		Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As DropDownBuilder Implements ILabelScreenTipSuperTip(Of DropDownBuilder).WithScreenTip
			_builder.WithScreentip(screenTip, callback)
			Return Me
		End Function

		Public Function WithSuperTip(superTip As String) As DropDownBuilder Implements ILabelScreenTipSuperTip(Of DropDownBuilder).WithSuperTip
			_builder.WithSupertip(superTip)
			Return Me
		End Function

		Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As DropDownBuilder Implements ILabelScreenTipSuperTip(Of DropDownBuilder).WithSuperTip
			_builder.WithSupertip(superTip, callback)
			Return Me
		End Function

		Public Function BeforeSelectionChange(ParamArray actions() As Action(Of DropDown.BeforeSelectionChangeEventArgs)) As DropDownBuilder
			For Each action As Action(Of DropDown.BeforeSelectionChangeEventArgs) In actions
				With New BeforeSelectionChangeEventHandlerHelper(action)
					_beforeSelectionChangeHandlers.Add(AddressOf .Handle)
				End With
			Next

			Return Me
		End Function

		Public Function AfterSelectionChanged(Paramarray actions() As Action(Of DropDown.SelectionChangedEventArgs)) As DropDownBuilder
			For Each action As Action(Of DropDown.SelectionChangedEventArgs) In actions
				With New SelectionChangeEventHandlerHelper(action)
					_selectionChangedHandlers.Add(AddressOf .Handle)
				End With
			Next

			Return Me
		End Function

		Private NotInheritable Class BeforeSelectionChangeEventHandlerHelper

			Private ReadOnly _action As Action(Of DropDown.BeforeSelectionChangeEventArgs)

			Public Sub New(action As Action(Of DropDown.BeforeSelectionChangeEventArgs))
				_action = action
			End Sub

			Public Sub Handle(sender As Object, e As DropDown.BeforeSelectionChangeEventArgs)
				_action.Invoke(e)
			End Sub

		End Class

		Private NotInheritable Class SelectionChangeEventHandlerHelper

			Private ReadOnly _action As Action(Of DropDown.SelectionChangedEventArgs)

			Public Sub New(action As Action(Of DropDown.SelectionChangedEventArgs))
				_action = action
			End Sub

			Public Sub Handle(sender As Object, e As DropDown.SelectionChangedEventArgs)
				_action.Invoke(e)
			End Sub

		End Class

	End Class

End Namespace