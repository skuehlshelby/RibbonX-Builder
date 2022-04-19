Imports RibbonFactory.BuilderInterfaces.API
Imports RibbonFactory.Containers
Imports RibbonFactory.Enums.MSO

Namespace Builders

    Friend NotInheritable Class ButtonGroupBuilder
        Inherits BuilderBase(Of ButtonGroup)
        Implements IButtonGroupBuilder

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(template As RibbonElement)
            MyBase.New(template)
        End Sub

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

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As IButtonGroupBuilder Implements IButtonGroupBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As IButtonGroupBuilder Implements IButtonGroupBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As IButtonGroupBuilder Implements IButtonGroupBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

    End Class

End Namespace