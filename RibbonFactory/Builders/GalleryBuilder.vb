Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Builders
	Public NotInheritable Class GalleryBuilder
		Implements IBuilder(Of Gallery)
		Implements IID(Of GalleryBuilder)
		Implements IEnabled(Of GalleryBuilder)
		Implements IVisible(Of GalleryBuilder)
		Implements IInsert(Of GalleryBuilder)
		Implements ILabelScreenTipSuperTip(Of GalleryBuilder)
		Implements IShowLabel(Of GalleryBuilder)
		Implements IDescription(Of GalleryBuilder)
		Implements IShowInRibbon(Of GalleryBuilder)
		Implements IInvalidateContentOnDrop(Of GalleryBuilder)
		Implements IImage(Of GalleryBuilder)
		Implements IShowImage(Of GalleryBuilder)
		Implements IKeyTip(Of GalleryBuilder)
		Implements ISize(Of GalleryBuilder)
		Implements IItemDimensions(Of GalleryBuilder)
		Implements IRowsAndColumns(Of GalleryBuilder)
		Implements IGetItemId(Of GalleryBuilder)
		Implements IGetItemCount(Of GalleryBuilder)
		Implements IGetItemLabel(Of GalleryBuilder)
		Implements IShowItemLabel(Of GalleryBuilder)
		Implements IGetItemScreentip(Of GalleryBuilder)
		Implements IGetItemSupertip(Of GalleryBuilder)
		Implements IGetItemImage(Of GalleryBuilder)
		Implements IShowItemImage(Of GalleryBuilder)
		Implements IGetSelectedItemId(Of GalleryBuilder)
		Implements IGetSelectedItemIndex(Of GalleryBuilder, Gallery)

		Private ReadOnly _builder As ControlBuilder
		Private ReadOnly _buttons As ICollection(Of Button) = New List(Of Button)

		Sub New()
			Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of Gallery)
			Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
			attributeGroupBuilder.SetDefaults(defaultProvider)
			_builder = New ControlBuilder(attributeGroupBuilder)
		End Sub

		Public Sub New(template As Gallery)
			Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
			attributeGroupBuilder.SetDefaults(template)
			attributeGroupBuilder.AddID(IdManager.GetID(Of Gallery)())
			_builder = New ControlBuilder(attributeGroupBuilder)
		End Sub

		Public Sub New(template As RibbonElement)
			Dim defaultProvider As IDefaultProvider = TryCast(template, IDefaultProvider)

			If defaultProvider IsNot Nothing Then
				Dim templateAttributes As AttributeSet = defaultProvider.GetDefaults()
				Dim galleryAttributes As AttributeSet = New DefaultProvider(Of Gallery)().GetDefaults()
				Dim intersection As AttributeSet = New AttributeSet(templateAttributes.Where(Function(a) galleryAttributes.Contains(a)))
				Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder(intersection)
				attributeGroupBuilder.AddID(IdManager.GetID(Of Gallery)())
				_builder = New ControlBuilder(attributeGroupBuilder)
			Else
				Throw New ArgumentException($"Could not copy applicable properties of type '{template.GetType().Name}' to type '{GetType(Gallery)}'")
			End If
		End Sub

		Public Function Build(Optional tag As Object = Nothing) As Gallery Implements IBuilder(Of Gallery).Build
			Return New Gallery(_builder.Build(), _buttons.ToArray(), tag:=tag)
		End Function

		Public Function WithId(id As String) As GalleryBuilder Implements IID(Of GalleryBuilder).WithId
			_builder.WithId(id)
			Return Me
		End Function

		Public Function WithIdQ([namespace] As String, id As String) As GalleryBuilder Implements IID(Of GalleryBuilder).WithIdQ
			_builder.WithId([namespace], id)
			Return Me
		End Function

		Public Function WithIdMso(mso As MSO) As GalleryBuilder Implements IID(Of GalleryBuilder).WithIdMso
			_builder.WithId(mso)
			Return Me
		End Function

		Public Function Enabled() As GalleryBuilder Implements IEnabled(Of GalleryBuilder).Enabled
			_builder.Enabled()
			Return Me
		End Function

		Public Function Enabled(callback As FromControl(Of Boolean)) As GalleryBuilder Implements IEnabled(Of GalleryBuilder).Enabled
			_builder.Enabled(callback)
			Return Me
		End Function

		Public Function Disabled() As GalleryBuilder Implements IEnabled(Of GalleryBuilder).Disabled
			_builder.Disabled()
			Return Me
		End Function

		Public Function Disabled(callback As FromControl(Of Boolean)) As GalleryBuilder Implements IEnabled(Of GalleryBuilder).Disabled
			_builder.Disabled(callback)
			Return Me
		End Function

		Public Function Visible() As GalleryBuilder Implements IVisible(Of GalleryBuilder).Visible
			_builder.Visible()
			Return Me
		End Function

		Public Function Visible(callback As FromControl(Of Boolean)) As GalleryBuilder Implements IVisible(Of GalleryBuilder).Visible
			_builder.Visible(callback)
			Return Me
		End Function

		Public Function Invisible() As GalleryBuilder Implements IVisible(Of GalleryBuilder).Invisible
			_builder.Invisible()
			Return Me
		End Function

		Public Function Invisible(callback As FromControl(Of Boolean)) As GalleryBuilder Implements IVisible(Of GalleryBuilder).Invisible
			_builder.Invisible(callback)
			Return Me
		End Function

		Public Function InsertBeforeMSO(builtInControl As MSO) As GalleryBuilder Implements IInsert(Of GalleryBuilder).InsertBefore
			_builder.InsertBefore(builtInControl)
			Return Me
		End Function

		Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As GalleryBuilder Implements IInsert(Of GalleryBuilder).InsertBefore
			_builder.InsertBefore(qualifiedControl)
			Return Me
		End Function

		Public Function InsertAfterMSO(builtInControl As MSO) As GalleryBuilder Implements IInsert(Of GalleryBuilder).InsertAfter
			_builder.InsertAfter(builtInControl)
			Return Me
		End Function

		Public Function InsertAfterQ(qualifiedControl As RibbonElement) As GalleryBuilder Implements IInsert(Of GalleryBuilder).InsertAfter
			_builder.InsertAfter(qualifiedControl)
			Return Me
		End Function

		Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As GalleryBuilder Implements ILabelScreenTipSuperTip(Of GalleryBuilder).WithLabel
			_builder.WithLabel(label, copyToScreenTip)
			_builder.ShowLabel()
			Return Me
		End Function

		Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As GalleryBuilder Implements ILabelScreenTipSuperTip(Of GalleryBuilder).WithLabel
			_builder.WithLabel(label, callback, copyToScreenTip)
			_builder.ShowLabel()
			Return Me
		End Function

		Public Function WithScreenTip(screenTip As String) As GalleryBuilder Implements ILabelScreenTipSuperTip(Of GalleryBuilder).WithScreenTip
			_builder.WithScreentip(screenTip)
			Return Me
		End Function

		Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As GalleryBuilder Implements ILabelScreenTipSuperTip(Of GalleryBuilder).WithScreenTip
			_builder.WithScreentip(screenTip, callback)
			Return Me
		End Function

		Public Function WithSuperTip(superTip As String) As GalleryBuilder Implements ILabelScreenTipSuperTip(Of GalleryBuilder).WithSuperTip
			_builder.WithSupertip(superTip)
			Return Me
		End Function

		Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As GalleryBuilder Implements ILabelScreenTipSuperTip(Of GalleryBuilder).WithSuperTip
			_builder.WithSupertip(superTip, callback)
			Return Me
		End Function

		Public Function ShowLabel() As GalleryBuilder Implements IShowLabel(Of GalleryBuilder).ShowLabel
			_builder.ShowLabel()
			Return Me
		End Function

		Public Function ShowLabel(callback As FromControl(Of Boolean)) As GalleryBuilder Implements IShowLabel(Of GalleryBuilder).ShowLabel
			_builder.ShowLabel(callback)
			Return Me
		End Function

		Public Function HideLabel() As GalleryBuilder Implements IShowLabel(Of GalleryBuilder).HideLabel
			_builder.HideLabel()
			Return Me
		End Function

		Public Function HideLabel(callback As FromControl(Of Boolean)) As GalleryBuilder Implements IShowLabel(Of GalleryBuilder).HideLabel
			_builder.HideLabel(callback)
			Return Me
		End Function

		Public Function InvalidateContentOnDrop() As GalleryBuilder Implements IInvalidateContentOnDrop(Of GalleryBuilder).InvalidateContentOnDrop
			_builder.InvalidateContentOnDrop()
			Return Me
		End Function

		Public Function WithImage(imagePath As String) As GalleryBuilder Implements IImage(Of GalleryBuilder).WithImage
			_builder.WithImage(imagePath)
			Return Me
		End Function

		Public Function WithImage(image As ImageMSO) As GalleryBuilder Implements IImage(Of GalleryBuilder).WithImage
			_builder.WithImage(image)
			Return Me
		End Function

		Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As GalleryBuilder Implements IImage(Of GalleryBuilder).WithImage
			_builder.WithImage(image, callback)
			Return Me
		End Function

		Public Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As GalleryBuilder Implements IImage(Of GalleryBuilder).WithImage
			_builder.WithImage(image, callback)
			Return Me
		End Function

		Public Function WithKeyTip(keyTip As KeyTip) As GalleryBuilder Implements IKeyTip(Of GalleryBuilder).WithKeyTip
			_builder.WithKeyTip(keyTip)
			Return Me
		End Function

		Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As GalleryBuilder Implements IKeyTip(Of GalleryBuilder).WithKeyTip
			_builder.WithKeyTip(keyTip, callback)
			Return Me
		End Function

		Public Function WithDescription(description As String) As GalleryBuilder Implements IDescription(Of GalleryBuilder).WithDescription
			_builder.WithDescription(description)
			Return Me
		End Function

		Public Function WithDescription(description As String, callback As FromControl(Of String)) As GalleryBuilder Implements IDescription(Of GalleryBuilder).WithDescription
			_builder.WithDescription(description, callback)
			Return Me
		End Function

		Public Function GetItemIdFrom(callback As FromControlAndIndex(Of String)) As GalleryBuilder Implements IGetItemId(Of GalleryBuilder).GetItemIdFrom
			_builder.GetItemIDFrom(callback)
			Return Me
		End Function

		Public Function GetItemCountFrom(callback As FromControl(Of Integer)) As GalleryBuilder Implements IGetItemCount(Of GalleryBuilder).GetItemCountFrom
			_builder.GetItemCountFrom(callback)
			Return Me
		End Function

		Public Function GetItemLabelFrom(callback As FromControlAndIndex(Of String)) As GalleryBuilder Implements IGetItemLabel(Of GalleryBuilder).GetItemLabelFrom
			_builder.GetItemLabelFrom(callback)
			Return Me
		End Function

		Public Function ShowLabelsOnChildItems() As GalleryBuilder Implements IShowItemLabel(Of GalleryBuilder).ShowLabelsOnChildItems
			_builder.ShowItemLabel()
			Return Me
		End Function

		Public Function HideLabelsOnChildItems() As GalleryBuilder Implements IShowItemLabel(Of GalleryBuilder).HideLabelsOnChildItems
			_builder.HideItemLabel()
			Return Me
		End Function

		Public Function GetItemScreenTipFrom(callback As FromControlAndIndex(Of String)) As GalleryBuilder Implements IGetItemScreentip(Of GalleryBuilder).GetItemScreenTipFrom
			_builder.GetItemScreentipFrom(callback)
			Return Me
		End Function

		Public Function GetItemSuperTipFrom(callback As FromControlAndIndex(Of String)) As GalleryBuilder Implements IGetItemSupertip(Of GalleryBuilder).GetItemSuperTipFrom
			_builder.GetItemSupertipFrom(callback)
			Return Me
		End Function

		Public Function GetItemImageFrom(callback As FromControlAndIndex(Of IPictureDisp)) As GalleryBuilder Implements IGetItemImage(Of GalleryBuilder).GetItemImageFrom
			_builder.GetItemImageFrom(callback)
			Return Me
		End Function

		Public Function ShowItemImages() As GalleryBuilder Implements IShowItemImage(Of GalleryBuilder).ShowItemImages
			_builder.ShowItemImages()
			Return Me
		End Function

		Public Function HideItemImages() As GalleryBuilder Implements IShowItemImage(Of GalleryBuilder).HideItemImages
			_builder.HideItemImages()
			Return Me
		End Function

		Public Function GetSelectedItemIdFrom(callback As FromControl(Of String), setSelected As SelectionChanged) As GalleryBuilder Implements IGetSelectedItemId(Of GalleryBuilder).GetSelectedItemIdFrom
			_builder.GetSelectedItemIDFrom(callback, setSelected)
			Return Me
		End Function

		Public Function GetSelectedItemIndexFrom(getSelected As FromControl(Of Integer), setSelected As SelectionChanged) As GalleryBuilder Implements IGetSelectedItemIndex(Of GalleryBuilder, Gallery).GetSelectedItemIndexFrom
			_builder.GetSelectedItemIndexFrom(getSelected, setSelected)
			Return Me
		End Function

		Public Function GetSelectedItemIndexFrom(getSelected As FromControl(Of Integer), setSelected As SelectionChanged, onSelectionChange As Action(Of Gallery)) As GalleryBuilder Implements IGetSelectedItemIndex(Of GalleryBuilder, Gallery).GetSelectedItemIndexFrom
			_builder.GetSelectedItemIndexFrom(getSelected, setSelected, onSelectionChange)
			Return Me
		End Function

		Public Function ShowImage() As GalleryBuilder Implements IShowImage(Of GalleryBuilder).ShowImage
			_builder.ShowImage()
			Return Me
		End Function

		Public Function ShowImage(getShowImage As FromControl(Of Boolean)) As GalleryBuilder Implements IShowImage(Of GalleryBuilder).ShowImage
			_builder.ShowImage(getShowImage)
			Return Me
		End Function

		Public Function HideImage() As GalleryBuilder Implements IShowImage(Of GalleryBuilder).HideImage
			_builder.HideImage()
			Return Me
		End Function

		Public Function HideImage(getShowImage As FromControl(Of Boolean)) As GalleryBuilder Implements IShowImage(Of GalleryBuilder).HideImage
			_builder.HideImage(getShowImage)
			Return Me
		End Function

		Public Function Large() As GalleryBuilder Implements ISize(Of GalleryBuilder).Large
			_builder.Large()
			Return Me
		End Function

		Public Function Large(callback As FromControl(Of Enums.ControlSize)) As GalleryBuilder Implements ISize(Of GalleryBuilder).Large
			_builder.Large(callback)
			Return Me
		End Function

		Public Function Normal() As GalleryBuilder Implements ISize(Of GalleryBuilder).Normal
			_builder.Normal()
			Return Me
		End Function

		Public Function Normal(callback As FromControl(Of Enums.ControlSize)) As GalleryBuilder Implements ISize(Of GalleryBuilder).Normal
			_builder.Normal(callback)
			Return Me
		End Function

		Public Function WithRowCount(rowCount As Integer) As GalleryBuilder Implements IRowsAndColumns(Of GalleryBuilder).WithRowCount
			_builder.WithRows(rowCount)
			Return Me
		End Function

		Public Function WithColumnCount(columnCount As Integer) As GalleryBuilder Implements IRowsAndColumns(Of GalleryBuilder).WithColumnCount
			_builder.WithColumns(columnCount)
			Return Me
		End Function

		Public Function ShowInRibbon() As GalleryBuilder Implements IShowInRibbon(Of GalleryBuilder).ShowInRibbon
			_builder.ShowInRibbon()
			Return Me
		End Function

		Public Function DoNotShowInRibbon() As GalleryBuilder Implements IShowInRibbon(Of GalleryBuilder).DoNotShowInRibbon
			_builder.DoNotShowInRibbon()
			Return Me
		End Function

		Public Function WithItemDimensions(dimensions As Size) As GalleryBuilder Implements IItemDimensions(Of GalleryBuilder).WithItemDimensions
			Return WithItemDimensions(dimensions.Height, dimensions.Width)
		End Function

		Public Function WithItemDimensions(height As Integer, width As Integer) As GalleryBuilder Implements IItemDimensions(Of GalleryBuilder).WithItemDimensions
			_builder.WithItemDimensions(height, width)
			Return Me
		End Function

		Public Function WithItemDimensions(dimensions As Size, heightCallback As FromControl(Of Integer), widthCallback As FromControl(Of Integer)) As GalleryBuilder Implements IItemDimensions(Of GalleryBuilder).WithItemDimensions
			Return WithItemDimensions(dimensions.Height, dimensions.Width, heightCallback, widthCallback)
		End Function

		Public Function WithItemDimensions(height As Integer, width As Integer, heightCallback As FromControl(Of Integer), widthCallback As FromControl(Of Integer)) As GalleryBuilder Implements IItemDimensions(Of GalleryBuilder).WithItemDimensions
			_builder.WithItemDimensions(height, width, heightCallback, widthCallback)
			Return Me
		End Function

	End Class

End Namespace