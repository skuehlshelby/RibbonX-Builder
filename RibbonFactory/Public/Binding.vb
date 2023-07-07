Imports System.Linq.Expressions

Namespace Api

    Public Interface IBinding
        Inherits IDisposable
    End Interface

    Public Interface IBindingBuilder
        Function OneWay(Of TSource As IRibbonElement, TTarget)(binding As Expression(Of Action(Of TSource, TTarget))) As IBinding
    End Interface

    Friend NotInheritable Class BindingBuilder
        Implements IBindingBuilder

        Public Function OneWay(Of TSource As IRibbonElement, TTarget)(binding As Expression(Of Action(Of TSource, TTarget))) As IBinding Implements IBindingBuilder.OneWay
            Throw New NotImplementedException()
        End Function
    End Class

    Friend NotInheritable Class Binding
        Implements IBinding

        Public Sub Dispose() Implements IDisposable.Dispose
            Throw New NotImplementedException()
        End Sub
    End Class

End Namespace