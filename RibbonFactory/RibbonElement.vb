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
Imports IRibbonUI = NetOffice.OfficeApi.IRibbonUI

Public MustInherit Class RibbonElement

#Region "Shared Members"
    Private Shared ReadOnly ControlIDs As Dictionary(Of Type, Integer) = New Dictionary(Of Type, Integer)
    Private Shared _ribbon As IRibbonUI

    Protected Shared Function FabricateID(Of T As RibbonElement)() As String
        Dim ribbonElementType As Type = GetType(T)

        If Not ControlIDs.ContainsKey(ribbonElementType) Then
            ControlIDs.Add(ribbonElementType, 0)
        End If

        ControlIDs.Item(ribbonElementType) += 1

        Return ribbonElementType.Name & ControlIDs.Item(ribbonElementType).ToString
    End Function

    Public Shared Sub AssignRibbonUI(ribbonUi As IRibbonUI)
        _ribbon = If(_ribbon, ribbonUi)
    End Sub

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
#End Region

#Region "Instance Members"
    Protected ReadOnly Attributes As AttributeGroup

    Protected Sub New(attributes As AttributeGroup, Optional tag As Object = Nothing)
        Me.Attributes = attributes
        Me.Tag = tag
    End Sub
    Public ReadOnly Property ID As String
        Get
            Return Attributes.Lookup(Of ID)().GetValue()
        End Get
    End Property

    Public Property Tag As Object

    Protected Sub Refresh()
        _ribbon.InvalidateControl(ID)
    End Sub

    Public MustOverride ReadOnly Property XML As String

    Public Overrides Function ToString() As String
        Return XML
    End Function

    Public MustOverride Overrides Function Equals(obj As Object) As Boolean

    Public Overrides Function GetHashCode() As Integer
        Return ID.GetHashCode()
    End Function
#End Region
End Class
