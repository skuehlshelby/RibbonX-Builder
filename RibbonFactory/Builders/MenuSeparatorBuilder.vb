Imports RibbonFactory.BuilderInterfaces.API
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums.MSO

Namespace Builders

    Friend NotInheritable Class MenuSeparatorBuilder
        Inherits BuilderBase(Of MenuSeparator)
        Implements IMenuSeparatorBuilder

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(template As RibbonElement)
            MyBase.New(template)
        End Sub

        Public Function WithId(id As String) As IMenuSeparatorBuilder Implements IMenuSeparatorBuilder.WithId
            WithIdBase(id)
            Return Me
        End Function

        Public Function WithIdQ([namespace] As String, id As String) As IMenuSeparatorBuilder Implements IMenuSeparatorBuilder.WithIdQ
            WithIdBase([namespace], id)
            Return Me
        End Function

        Public Function WithIdMso(mso As MSO) As IMenuSeparatorBuilder Implements IMenuSeparatorBuilder.WithIdMso
            WithIdBase(mso)
            Return Me
        End Function

        Public Function InsertBefore(builtInControl As MSO) As IMenuSeparatorBuilder Implements IMenuSeparatorBuilder.InsertBefore
            InsertBeforeMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertBefore(qualifiedControl As RibbonElement) As IMenuSeparatorBuilder Implements IMenuSeparatorBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfter(builtInControl As MSO) As IMenuSeparatorBuilder Implements IMenuSeparatorBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertAfter(qualifiedControl As RibbonElement) As IMenuSeparatorBuilder Implements IMenuSeparatorBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

        Public Function WithTitle(title As String) As IMenuSeparatorBuilder Implements IMenuSeparatorBuilder.WithTitle
            TitleBase(title)
            Return Me
        End Function

        Public Function WithTitle(title As String, callback As FromControl(Of String)) As IMenuSeparatorBuilder Implements IMenuSeparatorBuilder.WithTitle
            TitleBase(title, callback)
            Return Me
        End Function

    End Class

End Namespace