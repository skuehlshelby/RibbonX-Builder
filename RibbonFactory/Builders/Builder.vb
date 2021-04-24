Imports System.Drawing
Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.Description
Imports RibbonFactory.RibbonAttributes.Categories.Enabled
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.Insert
Imports RibbonFactory.RibbonAttributes.Categories.KeyTip
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.Screentip
Imports RibbonFactory.RibbonAttributes.Categories.Size
Imports RibbonFactory.RibbonAttributes.Categories.Supertip
Imports RibbonFactory.RibbonAttributes.Categories.Visible
Imports stdole

Namespace Builders
    Public MustInherit Class Builder
        Private Shared ReadOnly ControlIDs As Dictionary(Of Type, Integer) = New Dictionary(Of Type, Integer)

        Protected Shared Function FabricateID(Of T As RibbonElement)() As String
            Dim ribbonElementType As Type = GetType(T)

            If Not ControlIDs.ContainsKey(ribbonElementType) Then
                ControlIDs.Add(ribbonElementType, 0)
            End If

            ControlIDs.Item(ribbonElementType) += 1

            Return ribbonElementType.Name & ControlIDs.Item(ribbonElementType).ToString
        End Function

        Protected Shared Function GetDefaults(Of T As RibbonElement)() As AttributeGroup
            GetDefaults = New AttributeGroup From {
                New Id(FabricateID(Of T))
                }

            For Each interfaceType As Type In GetType(T).GetInterfaces
                Select Case interfaceType
                    Case GetType(IEnabled)
                        GetDefaults.Add(New Enabled(true))
                    Case GetType(IVisible)
                        GetDefaults.Add(New Visible(true))
                    Case GetType(ILabel)
                        GetDefaults.Add(New Label(String.Empty))
                    Case GetType(IShowLabel)
                        GetDefaults.Add(New ShowLabel(False))
                    Case GetType(IScreenTip)
                        GetDefaults.Add(New Screentip(String.Empty))
                    Case GetType(ISuperTip)
                        GetDefaults.Add(New Supertip(String.Empty))
                    Case GetType(ISize)
                        GetDefaults.Add(New RibbonAttributes.Categories.Size.Size(ControlSize.large))
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

        Protected ReadOnly Attributes As AttributeGroup = New AttributeGroup

        Public MustOverride Function Build() As T

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

        Protected Sub AddLabel(label As String)
            Attributes.Add(New Label(label))
        End Sub

        Protected Sub AddLabel(label As String, callback As FromControl(Of String))
            Attributes.Add(New GetLabel(label, callback))
        End Sub

        Protected Sub AddShowLabel(showLabel As Boolean)
            Attributes.Add(New ShowLabel(showLabel))
        End Sub

        Protected Sub AddScreenTip(screenTip As String)
            Attributes.Add(New Screentip(screenTip))
        End Sub

        Protected Sub AddScreenTip(screenTip As String, callback As FromControl(Of String))
            Attributes.Add(New GetScreentip(screenTip, callback))
        End Sub

        Protected Sub AddSuperTip(superTip As String)
            Attributes.Add(New Supertip(supertip))
        End Sub

        Protected Sub AddSuperTip(superTip As String, callback As FromControl(Of String))
            Attributes.Add(New GetSupertip(supertip, callback))
        End Sub

        Protected Sub AddDescription(description As String)
            Attributes.Add(new Description(description))
        End Sub

        Protected Sub AddDescription(description As String, callback As FromControl(Of String))
            Attributes.Add(new GetDescription(description, callback))
        End Sub

        Protected Sub AddAction(callback As OnAction, action As Action)
            Attributes.Add(New Categories.OnAction.OnAction(action, callback))
        End Sub
        
        Protected Sub InsertAfter(mso As MSO.MSO)
            Attributes.Add(New InsertAfterMso(mso.ToString()))
        End Sub
        
        Protected Sub InsertAfter(control As RibbonElement)
            Attributes.Add(New InsertAfterQ(control.ID))
        End Sub
        
        Protected Sub InsertBefore(mso As MSO.MSO)
            Attributes.Add(New InsertBeforeMso(mso.ToString()))
        End Sub
        
        Protected Sub InsertBefore(control As RibbonElement)
            Attributes.Add(New InsertBeforeQ(control.ID))
        End Sub
        
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
        
        Protected Sub AddKeyTip(keyTip As KeyTip)
            Attributes.Add(New Categories.KeyTip.KeyTip(keyTip))
        End Sub
        
        Protected Sub AddKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip))
            Attributes.Add(New GetKeyTip(keyTip, callback))
        End Sub

        Protected Sub AddImage(image As ImageMSO)
            Attributes.Add(New Categories.Image.ImageMso(image))
        End Sub

        Protected Sub AddImage(image As IPictureDisp, callback As FromControl(Of IPictureDisp)) 
            Attributes.Add(New Categories.Image.GetImage(image, callback))
        End Sub

        Protected Sub AddImage(image As Bitmap, callback As FromControl(Of IPictureDisp))
            Attributes.Add(New Categories.Image.GetImage(PictureDispConverter.ImageToPictureDisp(image), callback))
        End Sub
    End Class
End NameSpace