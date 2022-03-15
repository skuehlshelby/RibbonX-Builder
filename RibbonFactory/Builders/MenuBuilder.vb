Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.Utilities
Imports stdole

Namespace Builders
    Public NotInheritable Class MenuBuilder
        Implements IBuilder(Of Menu)
        Implements IInsert(Of MenuBuilder)
        Implements IEnabled(Of MenuBuilder)
        Implements IVisible(Of MenuBuilder)
        Implements ILabelScreenTipSuperTip(Of MenuBuilder)
        Implements IShowLabel(Of MenuBuilder)
        Implements IImage(Of MenuBuilder)
        Implements IItemSize(Of MenuBuilder)
        Implements IDescription(Of MenuBuilder)
        Implements ISize(Of MenuBuilder)
        Implements IKeyTip(Of MenuBuilder)

        Private ReadOnly _builder As ControlBuilder
        Private ReadOnly _controls As IList(Of RibbonElement)

        Public Sub New()
            Dim defaultProvider As IDefaultProvider = New Defaults(Of Menu)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = New ControlBuilder(attributeGroupBuilder)
            _controls = New List(Of RibbonElement)
        End Sub

        Public Sub New(template As Menu)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(template)
            attributeGroupBuilder.AddID(IdManager.GetID(Of Menu)())
            _builder = New ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Sub New(template As RibbonElement)
            Dim defaultProvider As IDefaultProvider = TryCast(template, IDefaultProvider)

            If defaultProvider IsNot Nothing Then
                Dim templateAttributes As AttributeSet = defaultProvider.GetDefaults()
                Dim menuAttributes As AttributeSet = New Defaults(Of Menu)().GetDefaults()
                Dim intersection As AttributeSet = New AttributeSet(templateAttributes.Where(Function(a) menuAttributes.Contains(a)))
                Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder(intersection)
                attributeGroupBuilder.AddID(IdManager.GetID(Of Menu)())
                _builder = New ControlBuilder(attributeGroupBuilder)
            Else
                Throw New ArgumentException($"Could not copy applicable properties of type '{template.GetType().Name}' to type '{GetType(Menu)}'")
            End If
        End Sub

        Public Function Build(Optional tag As Object = Nothing) As Menu Implements IBuilder(Of Menu).Build
            Return New Menu(_controls, _builder.Build(), tag)
        End Function

        Public Function WithControls(ParamArray buttons() As Button) As MenuBuilder
            Require(Of ArgumentNullException)(buttons.All(Function(b) b IsNot Nothing), $"Parameter '{NameOf(buttons)}', given to {NameOf(MenuBuilder)}, cannot be null or contain null values.")
            Array.ForEach(buttons, Sub(b) _controls.Add(b))
            Return Me
        End Function

        Public Function WithControls(ParamArray checkBoxes() As CheckBox) As MenuBuilder
            Require(Of ArgumentNullException)(checkBoxes.All(Function(c) c IsNot Nothing), $"Parameter '{NameOf(checkBoxes)}', given to {NameOf(MenuBuilder)}, cannot be null or contain null values.")
            Array.ForEach(checkBoxes, Sub(c) _controls.Add(c))
            Return Me
        End Function

        Public Function WithControls(ParamArray galleries() As Gallery) As MenuBuilder
            Require(Of ArgumentNullException)(galleries.All(Function(g) g IsNot Nothing), $"Parameter '{NameOf(galleries)}', given to {NameOf(MenuBuilder)}, cannot be null or contain null values.")
            Array.ForEach(galleries, Sub(g) _controls.Add(g))
            Return Me
        End Function

        Public Function WithControls(ParamArray menus() As Menu) As MenuBuilder
            Require(Of ArgumentNullException)(menus.All(Function(m) m IsNot Nothing), $"Parameter '{NameOf(menus)}', given to {NameOf(MenuBuilder)}, cannot be null or contain null values.")
            Array.ForEach(menus, Sub(m) _controls.Add(m))
            Return Me
        End Function

        Public Function WithControls(ParamArray menuSeparators() As MenuSeparator) As MenuBuilder
            Require(Of ArgumentNullException)(menuSeparators.All(Function(g) g IsNot Nothing), $"Parameter '{NameOf(menuSeparators)}', given to {NameOf(MenuBuilder)}, cannot be null or contain null values.")
            Array.ForEach(menuSeparators, Sub(g) _controls.Add(g))
            Return Me
        End Function

        Public Function WithControls(ParamArray splitButtons() As SplitButton) As MenuBuilder
            Require(Of ArgumentNullException)(splitButtons.All(Function(s) s IsNot Nothing), $"Parameter '{NameOf(splitButtons)}', given to {NameOf(MenuBuilder)}, cannot be null or contain null values.")
            Array.ForEach(splitButtons, Sub(s) _controls.Add(s))
            Return Me
        End Function

        Public Function WithControls(ParamArray toggleButtons() As ToggleButton) As MenuBuilder
            Require(Of ArgumentNullException)(toggleButtons.All(Function(t) t IsNot Nothing), $"Parameter '{NameOf(toggleButtons)}', given to {NameOf(MenuBuilder)}, cannot be null or contain null values.")
            Array.ForEach(toggleButtons, Sub(t) _controls.Add(t))
            Return Me
        End Function

        Public Function Enabled() As MenuBuilder Implements IEnabled(Of MenuBuilder).Enabled
            _builder.Enabled()
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As MenuBuilder Implements IEnabled(Of MenuBuilder).Enabled
            _builder.Enabled(callback)
            Return Me
        End Function

        Public Function Disabled() As MenuBuilder Implements IEnabled(Of MenuBuilder).Disabled
            _builder.Disabled()
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As MenuBuilder Implements IEnabled(Of MenuBuilder).Disabled
            _builder.Disabled(callback)
            Return Me
        End Function

        Public Function Visible() As MenuBuilder Implements IVisible(Of MenuBuilder).Visible
            _builder.Visible()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As MenuBuilder Implements IVisible(Of MenuBuilder).Visible
            _builder.Visible(callback)
            Return Me
        End Function

        Public Function Invisible() As MenuBuilder Implements IVisible(Of MenuBuilder).Invisible
            _builder.Invisible()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As MenuBuilder Implements IVisible(Of MenuBuilder).Invisible
            _builder.Invisible(callback)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As MenuBuilder Implements ILabelScreenTipSuperTip(Of MenuBuilder).WithLabel
            _builder.WithLabel(label, copyToScreenTip)
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As MenuBuilder Implements ILabelScreenTipSuperTip(Of MenuBuilder).WithLabel
            _builder.WithLabel(label, callback, copyToScreenTip)
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As MenuBuilder Implements ILabelScreenTipSuperTip(Of MenuBuilder).WithScreenTip
            _builder.WithScreentip(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As MenuBuilder Implements ILabelScreenTipSuperTip(Of MenuBuilder).WithScreenTip
            _builder.WithScreentip(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As MenuBuilder Implements ILabelScreenTipSuperTip(Of MenuBuilder).WithSuperTip
            _builder.WithSupertip(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As MenuBuilder Implements ILabelScreenTipSuperTip(Of MenuBuilder).WithSuperTip
            _builder.WithSupertip(superTip, callback)
            Return Me
        End Function

        Public Function ShowLabel() As MenuBuilder Implements IShowLabel(Of MenuBuilder).ShowLabel
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function ShowLabel(callback As FromControl(Of Boolean)) As MenuBuilder Implements IShowLabel(Of MenuBuilder).ShowLabel
            _builder.ShowLabel(callback)
            Return Me
        End Function

        Public Function HideLabel() As MenuBuilder Implements IShowLabel(Of MenuBuilder).HideLabel
            _builder.HideLabel()
            Return Me
        End Function

        Public Function HideLabel(callback As FromControl(Of Boolean)) As MenuBuilder Implements IShowLabel(Of MenuBuilder).HideLabel
            _builder.HideLabel(callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As MenuBuilder Implements IImage(Of MenuBuilder).WithImage
            _builder.WithImage(image)
            Return Me
        End Function

        Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As MenuBuilder Implements IImage(Of MenuBuilder).WithImage
            _builder.WithImage(image, callback)
            Return Me
        End Function

        Public Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As MenuBuilder Implements IImage(Of MenuBuilder).WithImage
            _builder.WithImage(image, callback)
            Return Me
        End Function

        Public Function WithImage(imagePath As String) As MenuBuilder Implements IImage(Of MenuBuilder).WithImage
            _builder.WithImage(imagePath)
            Return Me
        End Function

        Public Function WithDescription(description As String) As MenuBuilder Implements IDescription(Of MenuBuilder).WithDescription
            _builder.WithDescription(description)
            Return Me
        End Function

        Public Function WithDescription(description As String, callback As FromControl(Of String)) As MenuBuilder Implements IDescription(Of MenuBuilder).WithDescription
            _builder.WithDescription(description, callback)
            Return Me
        End Function

        Public Function Large() As MenuBuilder Implements ISize(Of MenuBuilder).Large
            _builder.Large()
            Return Me
        End Function

        Public Function Large(callback As FromControl(Of ControlSize)) As MenuBuilder Implements ISize(Of MenuBuilder).Large
            _builder.Large(callback)
            Return Me
        End Function

        Public Function Normal() As MenuBuilder Implements ISize(Of MenuBuilder).Normal
            _builder.Normal()
            Return Me
        End Function

        Public Function Normal(callback As FromControl(Of ControlSize)) As MenuBuilder Implements ISize(Of MenuBuilder).Normal
            _builder.Normal(callback)
            Return Me
        End Function

        Public Function WithLargeItems() As MenuBuilder Implements IItemSize(Of MenuBuilder).WithLargeItems
            _builder.WithLargeItems()
            Return Me
        End Function

        Public Function WithNormalSizeItems() As MenuBuilder Implements IItemSize(Of MenuBuilder).WithNormalSizeItems
            _builder.WithNormalSizeItems()
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As MenuBuilder Implements IKeyTip(Of MenuBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As MenuBuilder Implements IKeyTip(Of MenuBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip, callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As MenuBuilder Implements IInsert(Of MenuBuilder).InsertBefore
            _builder.InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As MenuBuilder Implements IInsert(Of MenuBuilder).InsertBefore
            _builder.InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As MenuBuilder Implements IInsert(Of MenuBuilder).InsertAfter
            _builder.InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As MenuBuilder Implements IInsert(Of MenuBuilder).InsertAfter
            _builder.InsertAfter(qualifiedControl)
            Return Me
        End Function

    End Class

End Namespace