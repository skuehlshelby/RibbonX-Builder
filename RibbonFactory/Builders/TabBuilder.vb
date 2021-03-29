Imports RibbonFactory.Builder_Interfaces
Imports RibbonFactory.Containers

Namespace Builders

    Public Class TabBuilder
        Inherits Builder(Of Tab)
        Implements ISetLabel(Of TabBuilder)
        Implements ISetVisibility(Of TabBuilder)

        Public Overrides Function Build() As Tab
            Throw New NotImplementedException()
        End Function

        Public Overrides Function Build(tag As Object) As Tab
            Throw New NotImplementedException()
        End Function

        Public Function Visible() As TabBuilder Implements ISetVisibility(Of TabBuilder).Visible
            Throw New NotImplementedException()
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As TabBuilder Implements ISetVisibility(Of TabBuilder).Visible
            Throw New NotImplementedException()
        End Function

        Public Function Invisible() As TabBuilder Implements ISetVisibility(Of TabBuilder).Invisible
            Throw New NotImplementedException()
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As TabBuilder Implements ISetVisibility(Of TabBuilder).Invisible
            Throw New NotImplementedException()
        End Function

        Public Function WithLabel(label As String) As TabBuilder Implements ISetLabel(Of TabBuilder).WithLabel
            Throw New NotImplementedException()
        End Function
    End Class

End NameSpace