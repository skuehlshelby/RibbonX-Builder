Imports System.Drawing
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Builders

    ''' <summary>
    ''' This class provides common methods for more specialized control builders to use.
    ''' More specialized control builders are facades which use a subset of the methods in this class.
    ''' This class, in combination with AttributeGroupBuilder, is designed to remove as much boilerplate
    ''' code as possible from the specialized control builder classes.
    ''' </summary>
    Friend NotInheritable Class ControlBuilder
        Private ReadOnly _attributeGroupBuilder As AttributeGroupBuilder

        Public Sub New(attributeGroupBuilder As AttributeGroupBuilder)
            _attributeGroupBuilder = attributeGroupBuilder
        End Sub

        Public Function Build() As AttributeSet
            Return _attributeGroupBuilder.Build()
        End Function

        Public Sub Enabled()
            _attributeGroupBuilder.AddEnabled(enabled:= True)
        End Sub

        Public Sub Enabled(callback As FromControl(Of Boolean))
            _attributeGroupBuilder.AddGetEnabled(enabled:= True, callback:= callback)
        End Sub

        Public Sub Disabled()
            _attributeGroupBuilder.AddEnabled(enabled:= False)
        End Sub

        Public Sub Disabled(callback As FromControl(Of Boolean))
            _attributeGroupBuilder.AddGetEnabled(enabled:= False, callback:= callback)
        End Sub

        Public Sub Visible()
            _attributeGroupBuilder.AddVisible(visible:= True)
        End Sub

        Public Sub Visible(callback As FromControl(Of Boolean))
            _attributeGroupBuilder.AddGetVisible(visible:= True, callback:= callback)
        End Sub

        Public Sub Invisible()
            _attributeGroupBuilder.AddVisible(visible:= False)
        End Sub

        Public Sub Invisible(callback As FromControl(Of Boolean))
            _attributeGroupBuilder.AddGetVisible(visible:= False, callback:= callback)
        End Sub

#Region "Positioning"

        Public Sub InsertAfter(mso As MSO)
            _attributeGroupBuilder.AddInsertAfterMso(mso:= mso)
        End Sub

        Public Sub InsertAfter(control As RibbonElement)
            _attributeGroupBuilder.AddInsertAfterQ(control:= control)
        End Sub

        Public Sub InsertBefore(mso As MSO)
            _attributeGroupBuilder.AddInsertBeforeMso(mso:= mso)
        End Sub

        Public Sub InsertBefore(control As RibbonElement)
            _attributeGroupBuilder.AddInsertBeforeQ(control:= control)
        End Sub

        Public Sub Horizontal()
            _attributeGroupBuilder.AddBoxStyle(style:= BoxStyle.horizontal)
        End Sub

        Public Sub Vertical()
            _attributeGroupBuilder.AddBoxStyle(style:= BoxStyle.vertical)
        End Sub

#End Region

#Region "Helptext"

        Public Sub WithLabel(label As String)
            _attributeGroupBuilder.AddLabel(label:= label)
        End Sub

        Public Sub WithLabel(label As String, copyToScreenTip As Boolean)
            _attributeGroupBuilder.AddLabel(label:= label)

            If copyToScreenTip Then _attributeGroupBuilder.AddScreentip(screentip:= label)
        End Sub

        Public Sub WithLabel(label As String, callback As FromControl(Of String))
            _attributeGroupBuilder.AddGetLabel(label:= label, callback:= callback)
        End Sub

        Public Sub WithLabel(label As String, callback As FromControl(Of String), copyToScreenTip As Boolean)
            _attributeGroupBuilder.AddGetLabel(label:= label, callback:= callback)

            If copyToScreenTip Then _attributeGroupBuilder.AddGetScreentip(screentip:= label, callback:= callback)
        End Sub

        Public Sub WithScreentip(screentip As String)
            _attributeGroupBuilder.AddScreentip(screentip:= screentip)
        End Sub

        Public Sub WithScreentip(screentip As String, callback As FromControl(Of String))
            _attributeGroupBuilder.AddGetScreentip(screentip:= screentip, callback:= callback)
        End Sub

        Public Sub WithSupertip(supertip As String)
            _attributeGroupBuilder.AddSupertip(supertip:= supertip)
        End Sub

        Public Sub WithSupertip(supertip As String, callback As FromControl(Of String))
            _attributeGroupBuilder.AddGetSupertip(supertip:= supertip, callback:= callback)
        End Sub

        Public Sub WithDescription(description As String)
            _attributeGroupBuilder.AddDescription(description:= description)
        End Sub

        Public Sub WithDescription(description As String, callback As FromControl(Of String))
            _attributeGroupBuilder.AddGetDescription(description:= description, callback:= callback)
        End Sub

        Public Sub ShowLabel()
            _attributeGroupBuilder.AddShowLabel(showLabel:= True)
        End Sub

        Public Sub ShowLabel(callback As FromControl(Of Boolean))
            _attributeGroupBuilder.AddGetShowLabel(showLabel:= True, callback:= callback)
        End Sub

        Public Sub HideLabel()
            _attributeGroupBuilder.AddShowLabel(showLabel:= False)
        End Sub

        Public Sub HideLabel(callback As FromControl(Of Boolean))
            _attributeGroupBuilder.AddGetShowLabel(showLabel:= False, callback:= callback)
        End Sub

        Public Sub WithKeyTip(keyTip As KeyTip)
            _attributeGroupBuilder.AddKeytip(keyTip:= keyTip)
        End Sub

        Public Sub WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip))
            _attributeGroupBuilder.AddGetKeytip(keyTip:= keyTip, callback:= callback)
        End Sub

#End Region

#Region "Images"

        Public Sub WithImage(imagePath As String)
            _attributeGroupBuilder.AddImage(path:= imagePath)
        End Sub

        Public Sub WithImage(imageMso As ImageMSO)
            _attributeGroupBuilder.AddImageMso(image:= imageMso)
        End Sub

        Public Sub WithImage(image As BitMap, callback As FromControl(Of IPictureDisp))
            _attributeGroupBuilder.AddGetImage(image:= ConvertToIPictureDisplay(image), callback:= callback)
        End Sub

        Public Sub WithImage(image As IPictureDisp, callback As FromControl(Of IPictureDisp))
            _attributeGroupBuilder.AddGetImage(image:= image, callback:= callback)
        End Sub

        Public Sub ShowImage()
            _attributeGroupBuilder.AddShowImage(showImage:= True)
        End Sub

        Public Sub ShowImage(callback As FromControl(Of Boolean))
            _attributeGroupBuilder.AddGetShowImage(showImage:= True, callback:= callback)
        End Sub

        Public Sub HideImage()
            _attributeGroupBuilder.AddShowImage(showImage:= False)
        End Sub

        Public Sub HideImage(callback As FromControl(Of Boolean))
            _attributeGroupBuilder.AddGetShowImage(showImage:= False, callback:= callback)
        End Sub

#End Region

#Region "Actions"

        Public Sub ThatDoes(action As Action, callback As OnAction)
            _attributeGroupBuilder.AddOnAction(action:= action, callback:= callback)
        End Sub

        Public Sub ThatDoes(action As Action, callback As ButtonPressed)
            _attributeGroupBuilder.AddOnAction(action:= action, callback:= callback)
        End Sub

        Public Sub ThatDoes(action As Action, callback As TextChanged)
            _attributeGroupBuilder.AddOnChange(action:= action, callback:= callback)
        End Sub

        Public Sub ThatDoes(action As Action, callback As SelectionChanged)
            _attributeGroupBuilder.AddOnAction(action:= action, callback:= callback)
        End Sub

#End Region

#Region "Content Fetching"

        Public Sub GetPressedFrom(callback As FromControl(Of Boolean))
            _attributeGroupBuilder.AddGetPressed(pressed:= False, callback:= callback)
        End Sub

        Public Sub GetTextFrom(callback As FromControl(Of String))
            _attributeGroupBuilder.AddGetText(text:= String.Empty, callback:= callback)
        End Sub

#End Region

#Region "Size"

        Public Sub Normal()
            _attributeGroupBuilder.AddSize(size:= ControlSize.normal)
        End Sub

        Public Sub Normal(callback As FromControl(Of ControlSize))
            _attributeGroupBuilder.AddGetSize(size:= ControlSize.normal, callback:= callback)
        End Sub

        Public Sub Large()
            _attributeGroupBuilder.AddSize(size:= ControlSize.large)
        End Sub

        Public Sub Large(callback As FromControl(Of ControlSize))
            _attributeGroupBuilder.AddGetSize(size:= ControlSize.large, callback:= callback)
        End Sub

        Public Sub WithSize(sizeString As String)
            _attributeGroupBuilder.AddSizeString(sizeString:= sizeString)
        End Sub

        Public Sub WithMaxLength(length As Byte)
            _attributeGroupBuilder.AddMaxLength(maxLength:= length)
        End Sub

#End Region

#Region "Child Items"

        Public Sub InvalidateContentOnDrop()
            _attributeGroupBuilder.AddInvalidateContentOnDrop(invalidateOnDrop:= True)
        End Sub

        Public Sub ShowItemImages()
            _attributeGroupBuilder.AddShowItemImage(showItemImage:= True)
        End Sub

        Public Sub HideItemImages()
            _attributeGroupBuilder.AddShowItemImage(showItemImage:= False)
        End Sub

        Public Sub ShowItemLabel()
            _attributeGroupBuilder.AddShowItemLabel(showItemLabel:= True)
        End Sub

        Public Sub HideItemLabel()
            _attributeGroupBuilder.AddShowItemLabel(showItemLabel:= False)
        End Sub

        Public Sub WithItemHeight(height As Integer)
            _attributeGroupBuilder.AddItemHeight(itemHeight:= height)
        End Sub

        Public Sub WithItemHeight(height As Integer, callback As FromControl(Of Integer))
            _attributeGroupBuilder.AddGetItemHeight(itemHeight:= height, callback:= callback)
        End Sub

        Public Sub WithItemWidth(width As Integer)
            _attributeGroupBuilder.AddItemWidth(itemWidth:= width)
        End Sub

        Public Sub WithItemWidth(width As Integer, callback As FromControl(Of Integer))
            _attributeGroupBuilder.AddGetItemWidth(itemWidth:= width, callback:= callback)
        End Sub

        Public Sub WithLargeItems()
            _attributeGroupBuilder.AddItemSize(ControlSize.large)
        End Sub

        Public Sub WithNormalSizeItems()
            _attributeGroupBuilder.AddItemSize(ControlSize.normal)
        End Sub

        Public Sub GetItemCountFrom(callback As FromControl(Of Integer))
            _attributeGroupBuilder.AddGetItemCount(callback:= callback)
        End Sub

        Public Sub GetItemIDFrom(callback As FromControlAndIndex(Of String))
            _attributeGroupBuilder.AddGetItemID(callback:= callback)
        End Sub

        Public Sub GetItemImageFrom(callback As FromControlAndIndex(Of IPictureDisp))
            _attributeGroupBuilder.AddGetItemImage(callback:= callback)
        End Sub

        Public Sub GetItemLabelFrom(callback As FromControlAndIndex(Of String))
            _attributeGroupBuilder.AddGetItemLabel(callback:= callback)
        End Sub

        Public Sub GetItemScreentipFrom(callback As FromControlAndIndex(Of String))
            _attributeGroupBuilder.AddGetItemScreentip(callback:= callback)
        End Sub

        Public Sub GetItemSupertipFrom(callback As FromControlAndIndex(Of String))
            _attributeGroupBuilder.AddGetItemSupertip(callback:= callback)
        End Sub

        Public Sub GetSelectedItemIDFrom(callback As FromControl(Of String))
            _attributeGroupBuilder.AddGetSelectedItemID(callback:= callback)
        End Sub

        Public Sub GetSelectedItemIndexFrom(callback As FromControl(Of Integer))
            _attributeGroupBuilder.AddGetSelectedItemIndex(callback:= callback)
        End Sub

        Public Sub WithRows(rowCount As Integer)
            _attributeGroupBuilder.AddRows(numberOfRows:= rowCount)
        End Sub

        Public Sub WithColumns(columnCount As Integer)
            _attributeGroupBuilder.AddColumns(numberOfColumns:= columnCount)
        End Sub

#End Region

#Region "Ribbon Only"

        Public Sub StartFromScratch()
            _attributeGroupBuilder.AddStartFromScratch(startFromScratch:= True)
        End Sub

        Public Sub OnLoad(callback As OnLoad)
            _attributeGroupBuilder.AddOnLoad(callback:= callback)
        End Sub

        Public Sub LoadImage(callback As LoadImage)
            _attributeGroupBuilder.AddLoadImage(callback:= callback)
        End Sub

#End Region

    End Class

End NameSpace