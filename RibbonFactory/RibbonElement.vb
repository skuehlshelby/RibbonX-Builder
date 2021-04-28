Imports NetOffice.OfficeApi

Public MustInherit Class RibbonElement
    Implements IEquatable(Of RibbonElement)
    
    Private Shared _ribbon As IRibbonUI

    Public Shared Sub SetRibbon(ribbon As IRibbonUI)
        _ribbon = If(_ribbon, ribbon)
    End Sub
    
    Protected Friend Sub New(Optional tag As Object = Nothing)
        Me.Tag = tag
    End Sub

    Public ReadOnly Property Ribbon As IRibbonUI
        Get
            Return _ribbon
        End Get
    End Property

    Protected Sub Refresh()
        Ribbon.InvalidateControl(ID)
    End Sub

    Public MustOverride ReadOnly Property ID As String

    Public Property Tag As Object

    Public MustOverride ReadOnly Property XML As String

    Public Overrides Function ToString() As String
        Return XML
    End Function

    Public  Overrides Function GetHashCode() As Integer
        Return ID.GetHashCode()
    End Function

    Public Overloads Overrides Function Equals(obj As Object) As Boolean
        Return Equals(TryCast(obj, RibbonElement))
    End Function
    
    Public Overloads Function Equals(other As RibbonElement) As Boolean Implements IEquatable(Of RibbonElement).Equals
        Return other IsNot Nothing AndAlso other.ID.Equals(ID)
    End Function
End Class
