Imports RibbonFactory.Builder_Interfaces
Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.Description
Imports RibbonFactory.RibbonAttributes.Categories.Enabled
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.Screentip
Imports RibbonFactory.RibbonAttributes.Categories.Size
Imports RibbonFactory.RibbonAttributes.Categories.Supertip
Imports RibbonFactory.RibbonAttributes.Categories.Visible

Namespace Builders
    Public MustInherit Class Builder
        Protected Shared ReadOnly ControlIDs As Dictionary(Of Type, Integer) = New Dictionary(Of Type, Integer)

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
                If interfaceType Is GetType(IEnable) Then
                    GetDefaults.Add(New Enabled(true))
                    GetDefaults.Add(New Visible(True))

                ElseIf interfaceType Is GetType(ITip) Then
                    GetDefaults.Add(New ShowLabel(False))
                    GetDefaults.Add(New Label(String.Empty))
                    GetDefaults.Add(New Screentip(String.Empty))
                    GetDefaults.Add(New Supertip(String.Empty))

                ElseIf interfaceType Is GetType(IEditable) Then

                ElseIf interfaceType Is GetType(IGraphic) Then
                    GetDefaults.Add(New Categories.Image.ImageMso(ImageMSO.HappyFace))

                ElseIf interfaceType Is GetType(IResizable) Then
                    GetDefaults.Add(New Size(ControlSize.large))

                ElseIf interfaceType Is GetType(ISelectable) Then

                ElseIf interfaceType Is GetType(IToggle) Then

                ElseIf interfaceType Is GetType(IDescribe) Then
                    GetDefaults.Add(New Description(String.Empty))

                End If
            Next interfaceType

        End Function
    End Class

    Public MustInherit Class Builder(Of T As RibbonElement)
        Inherits Builder
        Implements IBuilder(Of T)

        Public MustOverride Function Build() As T Implements IBuilder(Of T).Build

        
    End Class
End NameSpace