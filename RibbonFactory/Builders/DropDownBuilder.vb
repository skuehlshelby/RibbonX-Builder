Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports stdole

Namespace Builders

    Public Class DropDownBuilder
        Inherits Builder(Of DropDown)
        Implements ISetEnabled(Of DropDownBuilder)
        Implements ISetInsertionPoint(Of DropDownBuilder)
        Implements ISetVisibility(Of DropDownBuilder)
        Implements ISetLabelScreenTipAndSuperTip(Of DropDownBuilder)
        Implements ISetKeyTip(Of DropDownBuilder)
        Implements ISetLabelVisibility(Of DropDownBuilder)
        Implements ISetImage(Of DropDownBuilder)
        Implements ISetImageVisibility(Of DropDownBuilder)
        Implements ISetMaxLength(Of DropDownBuilder)
        Implements ISetOnSelectionChange(Of DropDownBuilder)
        Implements ISetChildItemIdSource(Of DropDownBuilder)
        Implements ISetChildItemCountSource(Of DropDownBuilder)
        Implements ISetChildItemLabelSource(Of DropDownBuilder)
        Implements ISetChildItemLabelVisibility(Of DropDownBuilder)
        Implements ISetChildItemScreenTipSource(Of DropDownBuilder)
        Implements ISetChildItemSuperTipSource(Of DropDownBuilder)
        Implements ISetChildItemImageSource(Of DropDownBuilder)
        Implements ISetChildItemImageVisibility(Of DropDownBuilder)
        Implements ISetSelectedItemIdSource(Of DropDownBuilder)
        Implements ISetSelectedItemIndexSource(Of DropDownBuilder)

        Public Overrides Function Build(tag As Object) As DropDown
            Return New DropDown(attributes:=Attributes, tag:=tag)
        End Function

        Public Function Enabled() As DropDownBuilder Implements ISetEnabled(Of DropDownBuilder).Enabled
            AddEnabled(enabled:=True)
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As DropDownBuilder Implements ISetEnabled(Of DropDownBuilder).Enabled
            AddEnabled(enabled:=True, callback:=callback)
            Return Me
        End Function

        Public Function Disabled() As DropDownBuilder Implements ISetEnabled(Of DropDownBuilder).Disabled
            AddEnabled(enabled:=False)
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As DropDownBuilder Implements ISetEnabled(Of DropDownBuilder).Disabled
            AddEnabled(enabled:=False, callback:=callback)
            Return Me
        End Function

        Public Function Visible() As DropDownBuilder Implements ISetVisibility(Of DropDownBuilder).Visible
            AddVisible(visible:=True)
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As DropDownBuilder Implements ISetVisibility(Of DropDownBuilder).Visible
            AddVisible(visible:=True, callback:=callback)
            Return Me
        End Function

        Public Function Invisible() As DropDownBuilder Implements ISetVisibility(Of DropDownBuilder).Invisible
            AddVisible(visible:=False)
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As DropDownBuilder Implements ISetVisibility(Of DropDownBuilder).Invisible
            AddVisible(visible:=False, callback:=callback)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As DropDownBuilder Implements ISetLabelScreenTipAndSuperTip(Of DropDownBuilder).WithLabel
            AddLabel(label:=label)
            AddShowLabel(showLabel:=True)

            If copyToScreenTip Then
                AddScreenTip(screenTip:=label)
            End If

            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As DropDownBuilder Implements ISetLabelScreenTipAndSuperTip(Of DropDownBuilder).WithLabel
            AddLabel(label:=label, callback:=callback)
            AddShowLabel(showLabel:=True)

            If copyToScreenTip Then
                AddScreenTip(screenTip:=label, callback:=callback)
            End If

            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As DropDownBuilder Implements ISetLabelScreenTipAndSuperTip(Of DropDownBuilder).WithScreenTip
            AddScreenTip(screenTip:=screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As DropDownBuilder Implements ISetLabelScreenTipAndSuperTip(Of DropDownBuilder).WithScreenTip
            AddScreenTip(screenTip:=screenTip, callback:=callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As DropDownBuilder Implements ISetLabelScreenTipAndSuperTip(Of DropDownBuilder).WithSuperTip
            AddSuperTip(superTip:=superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As DropDownBuilder Implements ISetLabelScreenTipAndSuperTip(Of DropDownBuilder).WithSuperTip
            AddSuperTip(superTip:=superTip, callback:=callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As DropDownBuilder Implements ISetKeyTip(Of DropDownBuilder).WithKeyTip
            AddKeyTip(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As DropDownBuilder Implements ISetKeyTip(Of DropDownBuilder).WithKeyTip
            AddKeyTip(keyTip, callback:=callback)
            Return Me
        End Function

        Public Function ShowLabel() As DropDownBuilder Implements ISetLabelVisibility(Of DropDownBuilder).ShowLabel
            AddShowLabel(showLabel:=True)
            Return Me
        End Function

        Public Function ShowLabel(callback As FromControl(Of Boolean)) As DropDownBuilder Implements ISetLabelVisibility(Of DropDownBuilder).ShowLabel
            AddShowLabel(showLabel:=True, callback:=callback)
            Return Me
        End Function

        Public Function HideLabel() As DropDownBuilder Implements ISetLabelVisibility(Of DropDownBuilder).HideLabel
            AddShowLabel(showLabel:=False)
            Return Me
        End Function

        Public Function HideLabel(callback As FromControl(Of Boolean)) As DropDownBuilder Implements ISetLabelVisibility(Of DropDownBuilder).HideLabel
            AddShowLabel(showLabel:=False, callback:=callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As DropDownBuilder Implements ISetImage(Of DropDownBuilder).WithImage
            AddImage(image)
            Return Me
        End Function

        Public Function WithImage(image As IPictureDisp, callback As FromControl(Of IPictureDisp)) As DropDownBuilder Implements ISetImage(Of DropDownBuilder).WithImage
            AddImage(image, callback:=callback)
            Return Me
        End Function

        Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As DropDownBuilder Implements ISetImage(Of DropDownBuilder).WithImage
            AddImage(image, callback:=callback)
            Return Me
        End Function

        Public Function ShowImage() As DropDownBuilder Implements ISetImageVisibility(Of DropDownBuilder).ShowImage
            AddShowImage(showImage:=True)
            Return Me
        End Function

        Public Function ShowImage(getShowImage As FromControl(Of Boolean)) As DropDownBuilder Implements ISetImageVisibility(Of DropDownBuilder).ShowImage
            AddShowImage(value:=True, callBack:=getShowImage)
            Return Me
        End Function

        Public Function HideImage() As DropDownBuilder Implements ISetImageVisibility(Of DropDownBuilder).HideImage
            AddShowImage(showImage:=False)
            Return Me
        End Function

        Public Function HideImage(getShowImage As FromControl(Of Boolean)) As DropDownBuilder Implements ISetImageVisibility(Of DropDownBuilder).HideImage
            AddShowImage(value:=False, callBack:=getShowImage)
            Return Me
        End Function

        Public Function OfWidth(limit As Byte) As DropDownBuilder Implements ISetMaxLength(Of DropDownBuilder).OfWidth
            AddMaxLength(limit)
            Return Me
        End Function

        Public Function OnSelectionChange(action As Action, callback As SelectionChanged) As DropDownBuilder Implements ISetOnSelectionChange(Of DropDownBuilder).OnSelectionChange
            AddAction(action:=action, callback:=callback)
            Return Me
        End Function

        Public Function GetItemIdFrom(idSource As FromControlAndIndex(Of String)) As DropDownBuilder Implements ISetChildItemIdSource(Of DropDownBuilder).GetItemIdFrom
            AddGetItemId(itemIdSource:=idSource)
            Return Me
        End Function

        Public Function GetItemCountFrom(itemCountSource As FromControl(Of Integer)) As DropDownBuilder Implements ISetChildItemCountSource(Of DropDownBuilder).GetItemCountFrom
            AddGetItemCount(itemCountSource:=itemCountSource)
            Return Me
        End Function

        Public Function GetItemLabelFrom(labelSource As FromControlAndIndex(Of String)) As DropDownBuilder Implements ISetChildItemLabelSource(Of DropDownBuilder).GetItemLabelFrom
            AddGetItemLabel(labelSource:=labelSource)
            Return Me
        End Function

        Public Function GetItemScreenTipFrom(screenTipSource As FromControlAndIndex(Of String)) As DropDownBuilder Implements ISetChildItemScreenTipSource(Of DropDownBuilder).GetItemScreenTipFrom
            AddGetItemScreenTip(screenTipSource:=screenTipSource)
            Return Me
        End Function

        Public Function GetItemSuperTipFrom(superTipSource As FromControlAndIndex(Of String)) As DropDownBuilder Implements ISetChildItemSuperTipSource(Of DropDownBuilder).GetItemSuperTipFrom
            AddGetItemSuperTip(superTipSource:=superTipSource)
            Return Me
        End Function

        Public Function GetItemImageFrom(imageSource As FromControlAndIndex(Of IPictureDisp)) As DropDownBuilder Implements ISetChildItemImageSource(Of DropDownBuilder).GetItemImageFrom
            AddGetItemImage(imageSource:=imageSource)
            Return Me
        End Function

        Public Function ShowImagesOnChildItems() As DropDownBuilder Implements ISetChildItemImageVisibility(Of DropDownBuilder).ShowImagesOnChildItems
            AddShowItemImage(showImage:=True)
            Return Me
        End Function

        Public Function HideImagesOnChildItems() As DropDownBuilder Implements ISetChildItemImageVisibility(Of DropDownBuilder).HideImagesOnChildItems
            AddShowItemImage(showImage:=False)
            Return Me
        End Function

        Public Function GetSelectedItemIdFrom(itemIdSource As FromControlAndIndex(Of String)) As DropDownBuilder Implements ISetSelectedItemIdSource(Of DropDownBuilder).GetSelectedItemIdFrom
            AddGetSelectedItemId(selectedItemIdSource:=itemIdSource)
            Return Me
        End Function

        Public Function GetSelectedItemIndexFrom(itemIndexSource As FromControl(Of Integer)) As DropDownBuilder Implements ISetSelectedItemIndexSource(Of DropDownBuilder).GetSelectedItemIndexFrom
            AddGetSelectedItemIndex(selectedItemIndexSource:=itemIndexSource)
            Return Me
        End Function

        Public Function ShowLabelsOnChildItems() As DropDownBuilder Implements ISetChildItemLabelVisibility(Of DropDownBuilder).ShowLabelsOnChildItems
            AddShowItemLabel(showLabel:=True)
            Return Me
        End Function

        Public Function HideLabelsOnChildItems() As DropDownBuilder Implements ISetChildItemLabelVisibility(Of DropDownBuilder).HideLabelsOnChildItems
            AddShowItemLabel(showLabel:=False)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As DropDownBuilder Implements ISetInsertionPoint(Of DropDownBuilder).InsertBefore
            InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As DropDownBuilder Implements ISetInsertionPoint(Of DropDownBuilder).InsertBefore
            InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As DropDownBuilder Implements ISetInsertionPoint(Of DropDownBuilder).InsertAfter
            InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As DropDownBuilder Implements ISetInsertionPoint(Of DropDownBuilder).InsertAfter
            InsertAfter(qualifiedControl)
            Return Me
        End Function

    End Class
End Namespace