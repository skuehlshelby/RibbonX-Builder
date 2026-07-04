Imports RibbonX.Api
Imports RibbonX.Api.Internal
Imports RibbonX.Callbacks
Imports RibbonX.Controls
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.InternalApi

Namespace Builders

    Friend NotInheritable Class SeparatorBuilder
        Inherits BuilderBase(Of Separator)
        Implements ISeparatorBuilder

        Public Function WithId(id As String) As ISeparatorBuilder Implements ISeparatorBuilder.WithId
            WithIdBase(id)
            Return Me
        End Function

        Public Function WithIdQ([namespace] As String, id As String) As ISeparatorBuilder Implements ISeparatorBuilder.WithIdQ
            WithIdBase([namespace], id)
            Return Me
        End Function

        Public Function WithIdMso(mso As MSO) As ISeparatorBuilder Implements ISeparatorBuilder.WithIdMso
            WithIdBase(mso)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As ISeparatorBuilder Implements ISeparatorBuilder.InsertBefore
            InsertBeforeMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As IRibbonElement) As ISeparatorBuilder Implements ISeparatorBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As ISeparatorBuilder Implements ISeparatorBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As IRibbonElement) As ISeparatorBuilder Implements ISeparatorBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

        Public Function Visible() As ISeparatorBuilder Implements ISeparatorBuilder.Visible
            VisibleBase()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As ISeparatorBuilder Implements ISeparatorBuilder.Visible
            VisibleBase(callback)
            Return Me
        End Function

        Public Function Invisible() As ISeparatorBuilder Implements ISeparatorBuilder.Invisible
            InvisibleBase()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As ISeparatorBuilder Implements ISeparatorBuilder.Invisible
            InvisibleBase(callback)
            Return Me
        End Function

        Public Function WithTag(tag As Object) As ISeparatorBuilder Implements ITag(Of ISeparatorBuilder).WithTag
            WithTagBase(tag)
            Return Me
        End Function

        Public Function FromTemplate(template As IRibbonElement) As ISeparatorBuilder Implements ITemplatable(Of ISeparatorBuilder).FromTemplate
            FromTemplateBase(template)
            Return Me
        End Function

        Public Overrides Function Build() As Separator
            Return New Separator(ControlProperties, Tag)
        End Function

        Public Shared Function FromAction(Optional action As Action(Of ISeparatorBuilder) = Nothing) As ISeparator
            Dim instance As SeparatorBuilder = New SeparatorBuilder()

            action?.Invoke(instance)

            Return instance.Build()
        End Function
    End Class

End Namespace

