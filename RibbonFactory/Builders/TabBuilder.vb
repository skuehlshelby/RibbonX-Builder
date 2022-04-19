Imports RibbonFactory.BuilderInterfaces.API
Imports RibbonFactory.Containers
Imports RibbonFactory.Enums.MSO

Namespace Builders

    Friend NotInheritable Class TabBuilder
        Inherits BuilderBase(Of Tab)
        Implements ITabBuilder

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(template As RibbonElement)
            MyBase.New(template)
        End Sub

        Public Function Visible() As ITabBuilder Implements ITabBuilder.Visible
            VisibleBase()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As ITabBuilder Implements ITabBuilder.Visible
            VisibleBase(callback)
            Return Me
        End Function

        Public Function Invisible() As ITabBuilder Implements ITabBuilder.Invisible
            InvisibleBase()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As ITabBuilder Implements ITabBuilder.Invisible
            InvisibleBase(callback)
            Return Me
        End Function

        Public Function WithLabel(label As String) As ITabBuilder Implements ITabBuilder.WithLabel
            LabelBase(label:=label)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As ITabBuilder Implements ITabBuilder.WithKeyTip
            KeyTipBase(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As ITabBuilder Implements ITabBuilder.WithKeyTip
            KeyTipBase(keyTip, callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As ITabBuilder Implements ITabBuilder.InsertBefore
            InsertBeforeMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As ITabBuilder Implements ITabBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As ITabBuilder Implements ITabBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As ITabBuilder Implements ITabBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

    End Class

End Namespace