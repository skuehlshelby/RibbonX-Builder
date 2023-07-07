Imports RibbonX.Api
Imports RibbonX.Api.Internal
Imports RibbonX.Callbacks
Imports RibbonX.Controls
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.InternalApi

Namespace Builders

    Friend NotInheritable Class MenuSeparatorBuilder
        Inherits BuilderBase(Of MenuSeparator)
        Implements IMenuSeparatorBuilder

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

        Public Function InsertBefore(qualifiedControl As IRibbonElement) As IMenuSeparatorBuilder Implements IMenuSeparatorBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfter(builtInControl As MSO) As IMenuSeparatorBuilder Implements IMenuSeparatorBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertAfter(qualifiedControl As IRibbonElement) As IMenuSeparatorBuilder Implements IMenuSeparatorBuilder.InsertAfter
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

        Public Function WithTag(tag As Object) As IMenuSeparatorBuilder Implements ITag(Of IMenuSeparatorBuilder).WithTag
            WithTagBase(tag)
            Return Me
        End Function

        Public Function FromTemplate(template As IRibbonElement) As IMenuSeparatorBuilder Implements ITemplatable(Of IMenuSeparatorBuilder).FromTemplate
            FromTemplateBase(template)
            Return Me
        End Function

        Public Overrides Function Build() As MenuSeparator
            Return New MenuSeparator(ControlProperties, Tag)
        End Function

        Public Shared Function FromAction(Optional action As Action(Of IMenuSeparatorBuilder) = Nothing) As IMenuSeparator
            Dim instance As MenuSeparatorBuilder = New MenuSeparatorBuilder()

            action?.Invoke(instance)

            Return instance.Build()
        End Function
    End Class

End Namespace