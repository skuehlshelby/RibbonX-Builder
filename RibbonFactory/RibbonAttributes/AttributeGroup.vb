Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes.Categories.Description
Imports RibbonFactory.RibbonAttributes.Categories.Enabled
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.Screentip
Imports RibbonFactory.RibbonAttributes.Categories.Size
Imports RibbonFactory.RibbonAttributes.Categories.Supertip
Imports RibbonFactory.RibbonAttributes.Categories.Visible

Namespace RibbonAttributes
    Public Class AttributeGroup
        Inherits HashSet(Of RibbonAttribute)

        Public Event ValueChanged

        Public Shadows Function Add(Of T As RibbonAttribute)(item As T) As T
            If Not MyBase.Add(item) Then
                Remove(item)
                MyBase.Add(item)
            End If

            Return item
        End Function

        Public Function Lookup(Of T As RibbonAttribute)() As T
            Return DirectCast(Me.Single(Function(attribute) TypeOf attribute Is T), T)
        End Function

        Public Function GetValue(Of TValue, TType As RibbonAttribute(Of TValue, TType)) As TValue
            Return DirectCast(Me.Single(Function(attribute) TypeOf attribute Is TType), TType).GetValue()
        End Function

        Public Sub SetValue(Of TValue, TType As RibbonAttribute(Of TValue, TType))(value As TValue)
            Dim currentValue As TValue

            With DirectCast(Me.Single(Function(attribute) TypeOf attribute Is TType), IMutableRibbonAttribute(Of TValue))
                currentValue = .GetValue()

                If Not currentValue.Equals(value) Then
                    .SetValue(value)
                    RaiseEvent ValueChanged
                End If
            End With
        End Sub

        Public Sub MergeWith(other As AttributeGroup)
            For Each attribute As RibbonAttribute In other
                Me.Add(attribute)
            Next attribute
        End Sub

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
            GetDefaults = New AttributeGroup()

            GetDefaults.Add(New Id(FabricateID(Of T)()))

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
                    GetDefaults.Add(New RibbonAttributes.Categories.Image.ImageMso(ImageMSO.HappyFace))

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
End Namespace
