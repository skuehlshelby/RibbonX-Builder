Imports RibbonX.Api
Imports RibbonX.Api.Internal
Imports RibbonX.Callbacks
Imports RibbonX.Controls
Imports RibbonX.Controls.BuiltIn

Namespace InternalApi

    Friend NotInheritable Class ButtonGroupBuilder
        Inherits BuilderBase(Of ButtonGroup)
        Implements IButtonGroupBuilder

        Private ReadOnly items As ISet(Of IRibbonElement) = New HashSet(Of IRibbonElement)()

        Public Function WithId(id As String) As IButtonGroupBuilder Implements IID(Of IButtonGroupBuilder).WithId
            WithIdBase(id)
            Return Me
        End Function

        Public Function WithIdQ([namespace] As String, id As String) As IButtonGroupBuilder Implements IID(Of IButtonGroupBuilder).WithIdQ
            WithIdBase([namespace], id)
            Return Me
        End Function

        Public Function WithIdMso(mso As MSO) As IButtonGroupBuilder Implements IID(Of IButtonGroupBuilder).WithIdMso
            WithIdBase(mso)
            Return Me
        End Function

        Public Function Visible() As IButtonGroupBuilder Implements IButtonGroupBuilder.Visible
            VisibleBase()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As IButtonGroupBuilder Implements IButtonGroupBuilder.Visible
            VisibleBase(callback)
            Return Me
        End Function

        Public Function Invisible() As IButtonGroupBuilder Implements IButtonGroupBuilder.Invisible
            InvisibleBase()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As IButtonGroupBuilder Implements IButtonGroupBuilder.Invisible
            InvisibleBase(callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As IButtonGroupBuilder Implements IButtonGroupBuilder.InsertBefore
            InsertBeforeMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As IRibbonElement) As IButtonGroupBuilder Implements IButtonGroupBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As IButtonGroupBuilder Implements IButtonGroupBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As IRibbonElement) As IButtonGroupBuilder Implements IButtonGroupBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

        Public Function WithTag(tag As Object) As IButtonGroupBuilder Implements IButtonGroupBuilder.WithTag
            WithTagBase(tag)
            Return Me
        End Function

        Public Function FromTemplate(template As IRibbonElement) As IButtonGroupBuilder Implements IButtonGroupBuilder.FromTemplate
            FromTemplateBase(template)
            Return Me
        End Function

        Public Function WithControls(ParamArray controls() As IButtonGroupAddable) As IButtonGroupBuilder Implements IButtonGroupBuilder.WithControls
            For Each control As IButtonGroupAddable In controls
                items.Add(control)
            Next
            Return Me
        End Function

        Public Overrides Function Build() As ButtonGroup
            Return New ButtonGroup(ControlProperties, items.ToArray(), Tag)
        End Function

        Public Shared Function FromAction(Optional action As Action(Of IButtonGroupBuilder) = Nothing) As IButtonGroup
            Dim instance As ButtonGroupBuilder = New ButtonGroupBuilder()

            action?.Invoke(instance)

            Return instance.Build()
        End Function

    End Class

End Namespace