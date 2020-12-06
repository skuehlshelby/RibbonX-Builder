Imports IRibbonUI = NetOffice.OfficeApi.IRibbonUI

#If DEBUG Then
Imports InternalsVisibleTo = System.Runtime.CompilerServices.InternalsVisibleToAttribute
<Assembly: InternalsVisibleTo("RibbonFactoryTests", AllInternalsVisible:=True)>
#End If

Public MustInherit Class RibbonElement

#Region "Shared Members"
    Private Shared ReadOnly ControlIDs As Dictionary(Of Type, Integer) = New Dictionary(Of Type, Integer)
    Private Shared Ribbon As IRibbonUI

    Public Shared Function FabricateID(Of T As RibbonElement)() As String
        Dim RibbonElementType As Type = GetType(T)

        If Not ControlIDs.ContainsKey(RibbonElementType) Then
            ControlIDs.Add(RibbonElementType, 0)
        End If

        ControlIDs.Item(RibbonElementType) += 1

        Return RibbonElementType.Name & ControlIDs.Item(RibbonElementType).ToString
    End Function

    Public Sub AssignRibbonUI(ByVal RibbonUI As IRibbonUI)
        Ribbon = If(Ribbon, RibbonUI)
    End Sub
#End Region

#Region "Instance Members"
    Private Structure TThis
        Public ID As String
        Public Tag As Object
    End Structure

    Private This As TThis
    Protected Sub New(ByVal ID As String, Optional ByVal Tag As Object = Nothing)
        This = New TThis With {
            .ID = ID,
            .Tag = Tag
        }
    End Sub
    Protected ReadOnly Property ID As String
        Get
            Return This.ID
        End Get
    End Property

    Public Property Tag As Object
        Get
            Return This.Tag
        End Get
        Set(value As Object)
            If This.Tag IsNot value Then
                This.Tag = value
            End If
        End Set
    End Property
    Protected Sub Refresh()
        Ribbon.InvalidateControl(This.ID)
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
