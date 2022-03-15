Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes

Namespace Builders
    Public NotInheritable Class SplitButtonBuilder
        Implements IBuilder(Of SplitButton)
        Implements IID(Of SplitButtonBuilder)
        Implements IInsert(Of SplitButtonBuilder)
        Implements IVisible(Of SplitButtonBuilder)
        Implements IEnabled(Of SplitButtonBuilder)
        Implements IKeyTip(Of SplitButtonBuilder)
        Implements IShowLabel(Of SplitButtonBuilder)
        Implements ISize(Of SplitButtonBuilder)

        Private ReadOnly _builder As ControlBuilder
        Private _button As RibbonElement
        Private _menu As Menu

        Public Sub New()
            Dim defaultProvider As IDefaultProvider = New Defaults(Of SplitButton)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = New ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Sub New(template As SplitButton)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(template)
            attributeGroupBuilder.AddID(IdManager.GetID(Of SplitButton)())
            _builder = New ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Sub New(template As RibbonElement)
            Dim defaultProvider As IDefaultProvider = TryCast(template, IDefaultProvider)

            If defaultProvider IsNot Nothing Then
                Dim templateAttributes As AttributeSet = defaultProvider.GetDefaults()
                Dim splitButtonAttributes As AttributeSet = New Defaults(Of SplitButton)().GetDefaults()
                Dim intersection As AttributeSet = New AttributeSet(templateAttributes.Where(Function(a) splitButtonAttributes.Contains(a)))
                Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder(intersection)
                attributeGroupBuilder.AddID(IdManager.GetID(Of SplitButton)())
                _builder = New ControlBuilder(attributeGroupBuilder)
            Else
                Throw New ArgumentException($"Could not copy applicable properties of type '{template.GetType().Name}' to type '{GetType(SplitButton)}'")
            End If
        End Sub

        Public Function Build(Optional tag As Object = Nothing) As SplitButton Implements IBuilder(Of SplitButton).Build
            Utilities.Require(Of InvalidOperationException)(
                If(_builder.IsBuiltInControl(),
                   _button Is Nothing AndAlso _menu Is Nothing,
                   _button IsNot Nothing AndAlso _menu IsNot Nothing),
                "When a control is a copy of a built-in control, no other properties may be specified. Otherwise, both a button and a menu are required.")

            If _builder.IsBuiltInControl() Then
                Return New SplitButton(_builder.Build())
            Else
                Return New SplitButton(button:=_button, menu:=_menu, attributes:=_builder.Build(), tag:=tag)
            End If
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

        Public Function WithId(id As String) As SplitButtonBuilder Implements IID(Of SplitButtonBuilder).WithId
            _builder.WithId(id)
            Return Me
        End Function

        Public Function WithIdQ([namespace] As String, id As String) As SplitButtonBuilder Implements IID(Of SplitButtonBuilder).WithIdQ
            _builder.WithId([namespace], id)
            Return Me
        End Function

        Public Function WithIdMso(mso As MSO) As SplitButtonBuilder Implements IID(Of SplitButtonBuilder).WithIdMso
            _builder.WithId(mso)
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