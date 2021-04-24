Imports RibbonFactory.Builder_Interfaces
Imports RibbonFactory.Containers

Namespace Builders

    Public Class TabBuilder
        Inherits Builder(Of Tab)
        Implements ISetLabel(Of TabBuilder)
        Implements ISetVisibility(Of TabBuilder)

        Public Overrides Function Build() As Tab
            Return Build(tag:= Nothing)
        End Function

        Public Overrides Function Build(tag As Object) As Tab
            Return New Tab(attributes:= Attributes, tag:= tag)
        End Function

        Public Function Visible() As TabBuilder Implements ISetVisibility(Of TabBuilder).Visible
            AddVisible(visible:= True)
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As TabBuilder Implements ISetVisibility(Of TabBuilder).Visible
            AddVisible(visible:= True, callback:= callback)
            Return Me
        End Function

        Public Function Invisible() As TabBuilder Implements ISetVisibility(Of TabBuilder).Invisible
            AddVisible(visible:= False)
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As TabBuilder Implements ISetVisibility(Of TabBuilder).Invisible
            AddVisible(visible:= False, callback:= callback)
            Return Me
        End Function

        Public Function WithLabel(label As String) As TabBuilder Implements ISetLabel(Of TabBuilder).WithLabel
            AddLabel(label:= label)
            Return Me
        End Function
    End Class

End NameSpace