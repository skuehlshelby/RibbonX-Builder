Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.Visible

Namespace Containers

    Public Class Box
        Inherits RibbonElement
        Implements IVisible
        Implements IContainer(Of RibbonElement)

        Private ReadOnly _orientation AS BoxStyle

        Public Sub New(buttonAttributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(buttonAttributes, tag)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Public ReadOnly Property Items As List(Of RibbonElement) Implements IContainer.Items
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return Attributes.Lookup(Of Visible).GetValue()
            End Get
            Set
                Attributes.Lookup(Of GetVisible).SetValue(value)
            End Set
        End Property

        Public Sub Add(ParamArray elements() As RibbonElement) Implements IContainer(Of RibbonElement).Add
            Throw New NotImplementedException()
        End Sub

        Public Sub Clear() Implements IContainer.Clear
            Throw New NotImplementedException()
        End Sub

        Public Overrides Function Equals(obj As Object) As Boolean
            Throw New NotImplementedException()
        End Function
    End Class
End NameSpace