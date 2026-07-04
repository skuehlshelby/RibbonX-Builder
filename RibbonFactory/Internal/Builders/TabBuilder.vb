Imports RibbonX.Api
Imports RibbonX.Api.Internal
Imports RibbonX.Callbacks
Imports RibbonX.Controls
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.InternalApi
Imports RibbonX.SimpleTypes

Namespace Builders

    Friend NotInheritable Class TabBuilder
        Inherits BuilderBase(Of Tab)
        Implements ITabBuilder

        Private ReadOnly groups As ISet(Of IGroup) = New HashSet(Of IGroup)()

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

        Public Function InsertBeforeQ(qualifiedControl As IRibbonElement) As ITabBuilder Implements ITabBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As ITabBuilder Implements ITabBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As IRibbonElement) As ITabBuilder Implements ITabBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

        Public Function WithTag(tag As Object) As ITabBuilder Implements ITag(Of ITabBuilder).WithTag
            WithTagBase(tag)
            Return Me
        End Function

        Public Function FromTemplate(template As IRibbonElement) As ITabBuilder Implements ITemplatable(Of ITabBuilder).FromTemplate
            FromTemplateBase(template)
            Return Me
        End Function

        Public Function WithGroups(ParamArray groups() As IGroup) As ITabBuilder Implements ITabBuilder.WithGroups
            Array.ForEach(groups, Sub(g) Me.groups.Add(g))
            Return Me
        End Function

        Public Overrides Function Build() As Tab
            Return New Tab(ControlProperties, groups.ToArray(), Tag)
        End Function

        Public Shared Function FromAction(Optional action As Action(Of ITabBuilder) = Nothing) As ITab
            Dim instance As TabBuilder = New TabBuilder()

            action?.Invoke(instance)

            Return instance.Build()
        End Function

    End Class

End Namespace