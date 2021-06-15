Imports System.Drawing
Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.ChildItems
Imports RibbonFactory.RibbonAttributes.Categories.ChildItems.Image
Imports RibbonFactory.RibbonAttributes.Categories.ChildItems.Label
Imports RibbonFactory.RibbonAttributes.Categories.ChildItems.ScreenTip
Imports RibbonFactory.RibbonAttributes.Categories.ChildItems.SuperTip
Imports RibbonFactory.RibbonAttributes.Categories.Description
Imports RibbonFactory.RibbonAttributes.Categories.Enabled
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.Image
Imports RibbonFactory.RibbonAttributes.Categories.Insert
Imports RibbonFactory.RibbonAttributes.Categories.KeyTip
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.Screentip
Imports RibbonFactory.RibbonAttributes.Categories.Size
Imports RibbonFactory.RibbonAttributes.Categories.SuperTip
Imports RibbonFactory.RibbonAttributes.Categories.Visible
Imports stdole

Namespace Builders
    Public MustInherit Class Builder
        Private Shared ReadOnly ControlIDs As Dictionary(Of Type, Integer) = New Dictionary(Of Type, Integer)

        Protected Shared Function FabricateId(Of T As RibbonElement)() As String
            Dim ribbonElementType As Type = GetType(T)

            If Not ControlIDs.ContainsKey(ribbonElementType) Then
                ControlIDs.Add(ribbonElementType, 0)
            End If

            ControlIDs.Item(ribbonElementType) += 1

            Return ribbonElementType.Name & ControlIDs.Item(ribbonElementType).ToString
        End Function

        Protected Shared Function GetDefaults(Of T As RibbonElement)() As AttributeGroup
            GetDefaults = New AttributeGroup From {
                New Id(FabricateId(Of T))
            }

            For Each interfaceType As Type In GetType(T).GetInterfaces
                Select Case interfaceType
                    Case GetType(IEnabled)
                        GetDefaults.Add(New Enabled(True))
                    Case GetType(IVisible)
                        GetDefaults.Add(New Visible(True))
                    Case GetType(ILabel)
                        GetDefaults.Add(New Label(String.Empty))
                    Case GetType(IShowLabel)
                        GetDefaults.Add(New ShowLabel(False))
                    Case GetType(IScreenTip)
                        GetDefaults.Add(New Screentip(String.Empty))
                    Case GetType(ISuperTip)
                        GetDefaults.Add(New Supertip(String.Empty))
                    Case GetType(ISize)
                        GetDefaults.Add(New Categories.Size.Size(ControlSize.large))
                    Case GetType(IDescription)
                        GetDefaults.Add(New Description(String.Empty))
                    Case GetType(IImage)
                        GetDefaults.Add(New Categories.Image.ImageMso(Common.HappyFace))
                End Select
            Next interfaceType

        End Function
    End Class

    Public MustInherit Class Builder(Of T As RibbonElement)
        Inherits Builder

        Protected ReadOnly Attributes As AttributeGroup = GetDefaults(Of T)()

        Public Overridable Function Build() As T
            Return Build(tag:=Nothing)
        End Function

        Public MustOverride Function Build(tag As Object) As T

        Protected Sub AddEnabled(enabled As Boolean)
            Attributes.Add(New Enabled(enabled))
        End Sub

        Protected Sub AddEnabled(enabled As Boolean, callback As FromControl(Of Boolean))
            Attributes.Add(New GetEnabled(enabled, callback))
        End Sub

        Protected Sub AddVisible(visible As Boolean)
            Attributes.Add(New Visible(visible))
        End Sub

        Protected Sub AddVisible(visible As Boolean, callback As FromControl(Of Boolean))
            Attributes.Add(New GetEnabled(visible, callback))
        End Sub

        Protected Sub InsertAfter(mso As MSO)
            Attributes.Add(New InsertAfterMso(mso.ToString()))
        End Sub

        Protected Sub InsertAfter(control As RibbonElement)
            Attributes.Add(New InsertAfterQ(control.ID))
        End Sub

        Protected Sub InsertBefore(mso As MSO)
            Attributes.Add(New InsertBeforeMso(mso.ToString()))
        End Sub

        Protected Sub InsertBefore(control As RibbonElement)
            Attributes.Add(New InsertBeforeQ(control.ID))
        End Sub

        Protected Sub AddKeyTip(keyTip As KeyTip)
            Attributes.Add(New Categories.KeyTip.Keytip(keyTip))
        End Sub

        Protected Sub AddKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip))
            Attributes.Add(New GetKeytip(keyTip, callback))
        End Sub

        Protected Sub AddOrientation(orientation As BoxStyle)
            Attributes.Add(New Categories.BoxStyle.BoxStyle(orientation))
        End Sub

#Region "Helptext Methods"

        Protected Sub AddLabel(label As String)
            Attributes.Add(New Label(label))
        End Sub

        Protected Sub AddLabel(label As String, callback As FromControl(Of String))
            Attributes.Add(New GetLabel(label, callback))
        End Sub

        Protected Sub AddShowLabel(showLabel As Boolean)
            Attributes.Add(New ShowLabel(showLabel))
        End Sub

        Protected Sub AddShowLabel(showLabel As Boolean, callback As FromControl(Of Boolean))
            Attributes.Add(New GetShowLabel(showLabel, callback))
        End Sub

        Protected Sub AddScreenTip(screenTip As String)
            Attributes.Add(New Screentip(screenTip))
        End Sub

        Protected Sub AddScreenTip(screenTip As String, callback As FromControl(Of String))
            Attributes.Add(New GetScreentip(screenTip, callback))
        End Sub

        Protected Sub AddSuperTip(superTip As String)
            Attributes.Add(New Supertip(superTip))
        End Sub

        Protected Sub AddSuperTip(superTip As String, callback As FromControl(Of String))
            Attributes.Add(New GetSuperTip(superTip, callback))
        End Sub

        Protected Sub AddDescription(description As String)
            Attributes.Add(New Description(description))
        End Sub

        Protected Sub AddDescription(description As String, callback As FromControl(Of String))
            Attributes.Add(New GetDescription(description, callback))
        End Sub

#End Region

#Region "Action Methods"

        Protected Sub AddAction(callback As OnAction, action As Action)
            Attributes.Add(New Categories.OnAction.OnAction(action, callback))
        End Sub

        Protected Sub AddAction(callback As ButtonPressed, action As Action)
            Attributes.Add(New Categories.OnAction.OnAction(action, callback))
        End Sub

        Protected Sub AddAction(callback As SelectionChanged, action As Action)
            Attributes.Add(New Categories.OnAction.OnAction(action, callback))
        End Sub

        Protected Sub AddAction(callback As TextChanged, action As Action)
            Attributes.Add(New Categories.OnChange.OnChange(action, callback))
        End Sub

#End Region

#Region "Image Methods"

        Protected Sub AddImage(image As Enums.ImageMSO.ImageMSO)
            Attributes.Add(New Categories.Image.ImageMso(image))
        End Sub

        Protected Sub AddImage(image As IPictureDisp, callback As FromControl(Of IPictureDisp))
            Attributes.Add(New GetImage(image, callback))
        End Sub

        Protected Sub AddImage(image As Bitmap, callback As FromControl(Of IPictureDisp))
            Attributes.Add(New GetImage(ImageToPictureDisp(image), callback))
        End Sub

        Protected Sub AddShowImage(showImage As Boolean)
            Attributes.Add(New ShowImage(value:=showImage))
        End Sub

        Protected Sub AddShowImage(value As Boolean, callBack As FromControl(Of Boolean))
            Attributes.Add(New GetShowImage(value:=value, callback:=callBack))
        End Sub

#End Region

#Region "Size Methods"

        Protected Sub AddLarge()
            Attributes.Add(New Categories.Size.Size(ControlSize.large))
        End Sub

        Protected Sub AddLarge(callback As FromControl(Of ControlSize))
            Attributes.Add(New GetSize(ControlSize.large, callback))
        End Sub

        Protected Sub AddNormal()
            Attributes.Add(New Categories.Size.Size(ControlSize.normal))
        End Sub

        Protected Sub AddNormal(callback As FromControl(Of ControlSize))
            Attributes.Add(New GetSize(ControlSize.normal, callback))
        End Sub

        Protected Sub AddMaxLength(maxLength As Byte)
            Attributes.Add(New Categories.MaxLength.MaxLength(maxLength))
        End Sub

#End Region

#Region "Child Item Methods"

        Protected Sub AddGetItemId(itemIdSource As FromControlAndIndex(Of String))
            Attributes.Add(New GetItemID(callback:=itemIdSource))
        End Sub

        Protected Sub AddGetItemCount(itemCountSource As FromControl(Of Integer))
            Attributes.Add(New GetItemCount(callback:=itemCountSource))
        End Sub

        Protected Sub AddGetItemLabel(labelSource As FromControlAndIndex(Of String))
            Attributes.Add(New GetItemLabel(callback:=labelSource))
        End Sub

        Protected Sub AddGetItemScreenTip(screenTipSource As FromControlAndIndex(Of String))
            Attributes.Add(New GetItemScreenTip(callback:=screenTipSource))
        End Sub

        Protected Sub AddGetItemSuperTip(superTipSource As FromControlAndIndex(Of String))
            Attributes.Add(New GetItemSuperTip(callback:=superTipSource))
        End Sub

        Protected Sub AddGetItemImage(imageSource As FromControlAndIndex(Of IPictureDisp))
            Attributes.Add(New GetItemImage(callback:=imageSource))
        End Sub

        Protected Sub AddShowItemImage(showImage As Boolean)
            Attributes.Add(New ShowItemImage(showImage))
        End Sub

        Protected Sub AddShowItemLabel(showLabel As Boolean)
            Attributes.Add(New ShowItemLabel(value:=showLabel))
        End Sub

        Protected Sub AddGetSelectedItemId(selectedItemIdSource As FromControlAndIndex(Of String))
            Attributes.Add(New GetItemID(callback:=selectedItemIdSource))
        End Sub

        Protected Sub AddGetSelectedItemIndex(selectedItemIndexSource As FromControl(Of Integer))
            Attributes.Add(New GetSelectedItemIndex(callback:=selectedItemIndexSource))
        End Sub

#End Region
    End Class
End NameSpace