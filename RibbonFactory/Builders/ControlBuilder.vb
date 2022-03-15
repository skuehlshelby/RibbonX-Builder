Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.Utilities
Imports stdole

Namespace Builders

	''' <summary>
	''' This class provides common methods for more specialized control builders to use.
	''' More specialized control builders are facades which use a subset of the methods in this class.
	''' This class, in combination with AttributeGroupBuilder, is designed to remove as much boilerplate
	''' code as possible from the specialized control builder classes.
	''' </summary>
	Friend Class ControlBuilder
		Private ReadOnly _attributeGroupBuilder As AttributeGroupBuilder

		Public Sub New(attributeGroupBuilder As AttributeGroupBuilder)
			_attributeGroupBuilder = attributeGroupBuilder
		End Sub

		Public Shared Function Init(Of T As RibbonElement)() As ControlBuilder
			Return New ControlBuilder(New AttributeGroupBuilder(New AttributeSet().WithDefaults(Of T)))
		End Function

		Public Shared Function Init(Of T As RibbonElement)(template As RibbonElement) As ControlBuilder
			If TypeOf template Is IDefaultProvider Then
                Return New ControlBuilder(New AttributeGroupBuilder(New AttributeSet().WithDefaults(Of T).OverwriteWithIntersectionOf(DirectCast(template, IDefaultProvider).GetDefaults())))
            Else
                Throw New ArgumentException($"Could not copy applicable properties of type '{template.GetType().Name}' to type '{GetType(Button)}'")
            End If
		End Function

		Public Function Build() As AttributeSet
			If IsBuiltInControl() Then
				Require(Of InvalidOperationException)(
					_attributeGroupBuilder.
					Except(_attributeGroupBuilder.
						Where(Function(a) a.Name = AttributeName.IdMso)).
					All(Function(a) a.
						GetType().
						IsDerivedFromGenericType(GetType(RibbonAttributeDefault(Of )))),
					"If the IdMso attribute is specified, no other attributes may be specified.")
			End If

			Return _attributeGroupBuilder.Build()
		End Function

		Public Sub Disabled()
			_attributeGroupBuilder.AddEnabled(enabled:=False)
		End Sub

		Public Sub Disabled(callback As FromControl(Of Boolean))
			_attributeGroupBuilder.AddGetEnabled(enabled:=False, callback:=callback)
		End Sub

		Public Sub Enabled()
			_attributeGroupBuilder.AddEnabled(enabled:=True)
		End Sub

		Public Sub Enabled(callback As FromControl(Of Boolean))
			_attributeGroupBuilder.AddGetEnabled(enabled:=True, callback:=callback)
		End Sub

		Public Sub Invisible()
			_attributeGroupBuilder.AddVisible(visible:=False)
		End Sub

		Public Sub Invisible(callback As FromControl(Of Boolean))
			_attributeGroupBuilder.AddGetVisible(visible:=False, callback:=callback)
		End Sub

		Public Function IsBuiltInControl() As Boolean
			Return _attributeGroupBuilder.Any(Function(a) a.Name = AttributeName.IdMso)
		End Function

		Public Sub Visible()
			_attributeGroupBuilder.AddVisible(visible:=True)
		End Sub

		Public Sub Visible(callback As FromControl(Of Boolean))
			_attributeGroupBuilder.AddGetVisible(visible:=True, callback:=callback)
		End Sub

		Public Sub WithId(id As String)
			_attributeGroupBuilder.AddID(id)
		End Sub

		Public Sub WithId([namespace] As String, id As String)
			_attributeGroupBuilder.AddIdQ([namespace], id)
		End Sub

		Public Sub WithId(mso As MSO)
			_attributeGroupBuilder.AddIdMso(mso)
		End Sub

#Region "Positioning"

		Public Sub Horizontal()
			_attributeGroupBuilder.AddBoxStyle(style:=BoxStyle.horizontal)
		End Sub

		Public Sub InsertAfter(mso As MSO)
			_attributeGroupBuilder.AddInsertAfterMso(mso:=mso)
		End Sub

		Public Sub InsertAfter(control As RibbonElement)
			_attributeGroupBuilder.AddInsertAfterQ(control:=control)
		End Sub

		Public Sub InsertBefore(mso As MSO)
			_attributeGroupBuilder.AddInsertBeforeMso(mso:=mso)
		End Sub

		Public Sub InsertBefore(control As RibbonElement)
			_attributeGroupBuilder.AddInsertBeforeQ(control:=control)
		End Sub

		Public Sub Vertical()
			_attributeGroupBuilder.AddBoxStyle(style:=BoxStyle.vertical)
		End Sub

#End Region

#Region "Helptext"

		Public Sub HideLabel()
			_attributeGroupBuilder.AddShowLabel(showLabel:=False)
		End Sub

		Public Sub HideLabel(callback As FromControl(Of Boolean))
			_attributeGroupBuilder.AddGetShowLabel(showLabel:=False, callback:=callback)
		End Sub

		Public Sub ShowLabel()
			_attributeGroupBuilder.AddShowLabel(showLabel:=True)
		End Sub

		Public Sub ShowLabel(callback As FromControl(Of Boolean))
			_attributeGroupBuilder.AddGetShowLabel(showLabel:=True, callback:=callback)
		End Sub

		Public Sub WithDescription(description As String)
			_attributeGroupBuilder.AddDescription(description:=description)
		End Sub

		Public Sub WithDescription(description As String, callback As FromControl(Of String))
			_attributeGroupBuilder.AddGetDescription(description:=description, callback:=callback)
		End Sub

		Public Sub WithKeyTip(keyTip As KeyTip)
			_attributeGroupBuilder.AddKeytip(keyTip:=keyTip)
		End Sub

		Public Sub WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip))
			_attributeGroupBuilder.AddGetKeytip(keyTip:=keyTip, callback:=callback)
		End Sub

		Public Sub WithLabel(label As String)
			_attributeGroupBuilder.AddLabel(label:=label)
		End Sub

		Public Sub WithLabel(label As String, copyToScreenTip As Boolean)
			_attributeGroupBuilder.AddLabel(label:=label)

			If copyToScreenTip Then _attributeGroupBuilder.AddScreentip(screentip:=label)
		End Sub

		Public Sub WithLabel(label As String, callback As FromControl(Of String))
			_attributeGroupBuilder.AddGetLabel(label:=label, callback:=callback)
		End Sub

		Public Sub WithLabel(label As String, callback As FromControl(Of String), copyToScreenTip As Boolean)
			_attributeGroupBuilder.AddGetLabel(label:=label, callback:=callback)

			If copyToScreenTip Then _attributeGroupBuilder.AddGetScreentip(screentip:=label, callback:=callback)
		End Sub

		Public Sub WithScreentip(screentip As String)
			_attributeGroupBuilder.AddScreentip(screentip:=screentip)
		End Sub

		Public Sub WithScreentip(screentip As String, callback As FromControl(Of String))
			_attributeGroupBuilder.AddGetScreentip(screentip:=screentip, callback:=callback)
		End Sub

		Public Sub WithSupertip(supertip As String)
			_attributeGroupBuilder.AddSupertip(supertip:=supertip)
		End Sub

		Public Sub WithSupertip(supertip As String, callback As FromControl(Of String))
			_attributeGroupBuilder.AddGetSupertip(supertip:=supertip, callback:=callback)
		End Sub

#End Region

#Region "Images"

		Public Sub HideImage()
			_attributeGroupBuilder.AddShowImage(showImage:=False)
		End Sub

		Public Sub HideImage(callback As FromControl(Of Boolean))
			_attributeGroupBuilder.AddGetShowImage(showImage:=False, callback:=callback)
		End Sub

		Public Sub ShowImage()
			_attributeGroupBuilder.AddShowImage(showImage:=True)
		End Sub

		Public Sub ShowImage(callback As FromControl(Of Boolean))
			_attributeGroupBuilder.AddGetShowImage(showImage:=True, callback:=callback)
		End Sub

		Public Sub WithImage(imagePath As String)
			_attributeGroupBuilder.AddImage(path:=imagePath)
		End Sub

		Public Sub WithImage(imageMso As ImageMSO)
			_attributeGroupBuilder.AddImageMso(image:=imageMso)
		End Sub

		Public Sub WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp))
			_attributeGroupBuilder.AddGetImage(image:=New AdapterForIPicture(image), callback:=callback)
		End Sub

		Public Sub WithImage(image As Icon, callback As FromControl(Of IPictureDisp))
			_attributeGroupBuilder.AddGetImage(image:=New AdapterForIPicture(image), callback:=callback)
		End Sub

#End Region

#Region "Actions"

		Public Sub ThatDoes(callback As OnAction, handlers As IEnumerable(Of EventHandler), Optional beforeEvent As Boolean = False)
			_attributeGroupBuilder.AddOnAction(callback:=callback, handlers:=handlers, beforeEvent:=beforeEvent)
		End Sub

		Public Sub ThatDoes(Of T)(callback As OnAction, handlers As IEnumerable(Of EventHandler(Of T)), Optional beforeEvent As Boolean = False)
			_attributeGroupBuilder.AddOnAction(callback:=callback, handlers:=handlers, beforeEvent:=beforeEvent)
		End Sub

		Public Function AsEventHandler(Of TElement As RibbonElement)(action As Action(Of TElement)) As EventHandler
            Return AddressOf New ActionToEventHandlerAdapter(Of TElement)(action).AsEventHandler
        End Function

		Public Function AsEventHandler(Of TElement As RibbonElement, TArgs As EventArgs)(action As Action(Of TElement, TArgs)) As EventHandler(Of TArgs)
            Return AddressOf New ActionToEventHandlerAdapter(Of TElement, TArgs)(action).AsEventHandler
        End Function

		Private NotInheritable Class ActionToEventHandlerAdapter(Of TElement As RibbonElement)
			Private ReadOnly value As Action(Of TElement)

			Public Sub New(value As Action(Of TElement))
				Me.value = value
			End Sub

			Public Sub AsEventHandler(sender As Object, e As EventArgs)
				value.Invoke(DirectCast(sender, TElement))
			End Sub
		End Class

		Private NotInheritable Class ActionToEventHandlerAdapter(Of TElement As RibbonElement, TArgs As EventArgs)
			Private ReadOnly value As Action(Of TElement, TArgs)

			Public Sub New(value As Action(Of TElement, TArgs))
				Me.value = value
			End Sub

			Public Sub AsEventHandler(sender As Object, e As TArgs)
				value.Invoke(DirectCast(sender, TElement), e)
			End Sub
		End Class

#End Region

#Region "Updatable From UI"

		Public Sub GetPressedFrom(pressed As Boolean, getPressed As FromControl(Of Boolean), setPressed As ButtonPressed)
			_attributeGroupBuilder.AddGetPressed(pressed:=pressed, getPressed:=getPressed, setPressed:=setPressed)
		End Sub

		Public Sub GetSelectedItemIDFrom(getSelected As FromControl(Of String), setSelected As SelectionChanged)
			_attributeGroupBuilder.AddGetSelected(getSelected, setSelected)
		End Sub

		Public Sub GetSelectedItemIndexFrom(getSelected As FromControl(Of Integer), setSelected As SelectionChanged)
			_attributeGroupBuilder.AddGetSelected(getSelected, setSelected)
		End Sub

		Public Sub GetTextFrom(text As String, getText As FromControl(Of String), setText As TextChanged)
			_attributeGroupBuilder.AddGetText(text:=text, getText:=getText, setText:=setText)
		End Sub

#End Region

#Region "Size"

		Public Sub Large()
			_attributeGroupBuilder.AddSize(size:=ControlSize.large)
		End Sub

		Public Sub Large(callback As FromControl(Of ControlSize))
			_attributeGroupBuilder.AddGetSize(size:=ControlSize.large, callback:=callback)
		End Sub

		Public Sub Normal()
			_attributeGroupBuilder.AddSize(size:=ControlSize.normal)
		End Sub

		Public Sub Normal(callback As FromControl(Of ControlSize))
			_attributeGroupBuilder.AddGetSize(size:=ControlSize.normal, callback:=callback)
		End Sub

		Public Sub WithMaxLength(length As Byte)
			_attributeGroupBuilder.AddMaxLength(maxLength:=length)
		End Sub

		Public Sub WithSize(sizeString As String)
			_attributeGroupBuilder.AddSizeString(sizeString:=sizeString)
		End Sub

#End Region

#Region "Child Items"

		Public Sub GetItemCountFrom(callback As FromControl(Of Integer))
			_attributeGroupBuilder.AddGetItemCount(callback:=callback)
		End Sub

		Public Sub GetItemIDFrom(callback As FromControlAndIndex(Of String))
			_attributeGroupBuilder.AddGetItemID(callback:=callback)
		End Sub

		Public Sub GetItemImageFrom(callback As FromControlAndIndex(Of IPictureDisp))
			_attributeGroupBuilder.AddGetItemImage(callback:=callback)
		End Sub

		Public Sub GetItemLabelFrom(callback As FromControlAndIndex(Of String))
			_attributeGroupBuilder.AddGetItemLabel(callback:=callback)
		End Sub

		Public Sub GetItemScreentipFrom(callback As FromControlAndIndex(Of String))
			_attributeGroupBuilder.AddGetItemScreentip(callback:=callback)
		End Sub

		Public Sub GetItemSupertipFrom(callback As FromControlAndIndex(Of String))
			_attributeGroupBuilder.AddGetItemSupertip(callback:=callback)
		End Sub

		Public Sub HideItemImages()
			_attributeGroupBuilder.AddShowItemImage(showItemImage:=False)
		End Sub

		Public Sub HideItemLabel()
			_attributeGroupBuilder.AddShowItemLabel(showItemLabel:=False)
		End Sub

		Public Sub InvalidateContentOnDrop()
			_attributeGroupBuilder.AddInvalidateContentOnDrop(invalidateOnDrop:=True)
		End Sub

		Public Sub ShowItemImages()
			_attributeGroupBuilder.AddShowItemImage(showItemImage:=True)
		End Sub

		Public Sub ShowItemLabel()
			_attributeGroupBuilder.AddShowItemLabel(showItemLabel:=True)
		End Sub

		Public Sub WithColumns(columnCount As Integer)
			_attributeGroupBuilder.AddColumns(numberOfColumns:=columnCount)
		End Sub

		Public Sub WithItemHeight(height As Integer)
			_attributeGroupBuilder.AddItemHeight(itemHeight:=height)
		End Sub

		Public Sub WithItemHeight(height As Integer, callback As FromControl(Of Integer))
			_attributeGroupBuilder.AddGetItemHeight(itemHeight:=height, callback:=callback)
		End Sub

		Public Sub WithItemWidth(width As Integer)
			_attributeGroupBuilder.AddItemWidth(itemWidth:=width)
		End Sub

		Public Sub WithItemWidth(width As Integer, callback As FromControl(Of Integer))
			_attributeGroupBuilder.AddGetItemWidth(itemWidth:=width, callback:=callback)
		End Sub

		Public Sub WithLargeItems()
			_attributeGroupBuilder.AddItemSize(ControlSize.large)
		End Sub

		Public Sub WithNormalSizeItems()
			_attributeGroupBuilder.AddItemSize(ControlSize.normal)
		End Sub

		Public Sub WithRows(rowCount As Integer)
			_attributeGroupBuilder.AddRows(numberOfRows:=rowCount)
		End Sub

#End Region

#Region "Ribbon Only"

		Public Sub LoadImage(callback As LoadImage)
			_attributeGroupBuilder.AddLoadImage(callback:=callback)
		End Sub

		Public Sub OnLoad(callback As OnLoad)
			_attributeGroupBuilder.AddOnLoad(callback:=callback)
		End Sub

		Public Sub StartFromScratch()
			_attributeGroupBuilder.AddStartFromScratch(startFromScratch:=True)
		End Sub

#End Region

#Region "Gallery Only"

		Public Sub DoNotShowInRibbon()
			_attributeGroupBuilder.AddShowInRibbon(showInRibbon:=False)
		End Sub

		Public Sub ShowInRibbon()
			_attributeGroupBuilder.AddShowInRibbon(showInRibbon:=True)
		End Sub

		Public Sub WithItemDimensions(height As Integer, width As Integer)
			_attributeGroupBuilder.AddItemHeight(height)
			_attributeGroupBuilder.AddItemWidth(width)
		End Sub

		Public Sub WithItemDimensions(height As Integer, width As Integer, heightCallback As FromControl(Of Integer), widthCallback As FromControl(Of Integer))
			_attributeGroupBuilder.AddGetItemHeight(height, heightCallback)
			_attributeGroupBuilder.AddGetItemWidth(width, widthCallback)
		End Sub

#End Region

	End Class

End Namespace