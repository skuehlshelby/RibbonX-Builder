Imports RibbonX.BuilderInterfaces.API
Imports RibbonX.Controls
Imports RibbonX.Enums.MSO

Namespace Builders

    Friend NotInheritable Class SeparatorBuilder
        Inherits BuilderBase(Of Separator)
        Implements ISeparatorBuilder

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(template As RibbonElement)
            MyBase.New(template)
        End Sub

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

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As ISeparatorBuilder Implements ISeparatorBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As ISeparatorBuilder Implements ISeparatorBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As ISeparatorBuilder Implements ISeparatorBuilder.InsertAfter
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

    End Class

End Namespace

