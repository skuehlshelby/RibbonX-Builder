Imports System.Diagnostics.Contracts
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes

Namespace Builders
    Public NotInheritable Class SplitButtonBuilder
        Implements IInsert(Of SplitButtonBuilder)
        Implements IVisible(Of SplitButtonBuilder)
        Implements IEnabled(Of SplitButtonBuilder)
        Implements IKeyTip(Of SplitButtonBuilder)
        Implements IShowLabel(Of SplitButtonBuilder)
        Implements ISize(Of SplitButtonBuilder)

        Private ReadOnly _builder As ControlBuilder
        Private _button As RibbonElement
        Private _menu As Menu
        Friend Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of SplitButton)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = new ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Function Build(Optional tag As Object = Nothing) As SplitButton
            Return New SplitButton(button:=_button, menu:=_menu, attributes:=_builder.Build(), tag:=tag)
        End Function

        Public Function WithButton(button As Button) As SplitButtonBuilder
            _button = button
            Return Me
        End Function

        Public Function WithToggleButton(toggleButton As ToggleButton) As SplitButtonBuilder
            _button = toggleButton
            Return Me
        End Function

        Public Function WithMenu(menu As Menu) As SplitButtonBuilder
            _menu = menu
            Return Me
        End Function

        Public Function Enabled() As SplitButtonBuilder Implements IEnabled(Of SplitButtonBuilder).Enabled
            _builder.Enabled()
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As SplitButtonBuilder Implements IEnabled(Of SplitButtonBuilder).Enabled
            _builder.Enabled(callback)
            Return Me
        End Function

        Public Function Disabled() As SplitButtonBuilder Implements IEnabled(Of SplitButtonBuilder).Disabled
            _builder.Disabled()
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As SplitButtonBuilder Implements IEnabled(Of SplitButtonBuilder).Disabled
            _builder.Disabled(callback)
            Return Me
        End Function

        Public Function Visible() As SplitButtonBuilder Implements IVisible(Of SplitButtonBuilder).Visible
            _builder.Visible()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As SplitButtonBuilder Implements IVisible(Of SplitButtonBuilder).Visible
            _builder.Visible(callback)
            Return Me
        End Function

        Public Function Invisible() As SplitButtonBuilder Implements IVisible(Of SplitButtonBuilder).Invisible
            _builder.Invisible()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As SplitButtonBuilder Implements IVisible(Of SplitButtonBuilder).Invisible
            _builder.Invisible(callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As SplitButtonBuilder Implements IKeyTip(Of SplitButtonBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As SplitButtonBuilder Implements IKeyTip(Of SplitButtonBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip, callback)
            Return Me
        End Function

        Public Function ShowLabel() As SplitButtonBuilder Implements IShowLabel(Of SplitButtonBuilder).ShowLabel
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function ShowLabel(callback As FromControl(Of Boolean)) As SplitButtonBuilder Implements IShowLabel(Of SplitButtonBuilder).ShowLabel
            _builder.ShowLabel(callback)
            Return Me
        End Function

        Public Function HideLabel() As SplitButtonBuilder Implements IShowLabel(Of SplitButtonBuilder).HideLabel
            _builder.HideLabel()
            Return Me
        End Function

        Public Function HideLabel(callback As FromControl(Of Boolean)) As SplitButtonBuilder Implements IShowLabel(Of SplitButtonBuilder).HideLabel
            _builder.HideLabel(callback)
            Return Me
        End Function

        Public Function Large() As SplitButtonBuilder Implements ISize(Of SplitButtonBuilder).Large
            _builder.Large()
            Return Me
        End Function

        Public Function Large(callback As FromControl(Of ControlSize)) As SplitButtonBuilder Implements ISize(Of SplitButtonBuilder).Large
            _builder.Large(callback)
            Return Me
        End Function

        Public Function Normal() As SplitButtonBuilder Implements ISize(Of SplitButtonBuilder).Normal
            _builder.Normal()
            Return Me
        End Function

        Public Function Normal(callback As FromControl(Of ControlSize)) As SplitButtonBuilder Implements ISize(Of SplitButtonBuilder).Normal
            _builder.Normal(callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As SplitButtonBuilder Implements IInsert(Of SplitButtonBuilder).InsertBefore
            _builder.InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As SplitButtonBuilder Implements IInsert(Of SplitButtonBuilder).InsertBefore
            _builder.InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As SplitButtonBuilder Implements IInsert(Of SplitButtonBuilder).InsertAfter
            _builder.InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As SplitButtonBuilder Implements IInsert(Of SplitButtonBuilder).InsertAfter
            _builder.InsertAfter(qualifiedControl)
            Return Me
        End Function

    End Class

End Namespace