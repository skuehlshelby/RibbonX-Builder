Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes

Namespace Builders

	Public NotInheritable Class ButtonGroupBuilder
		Implements IInsert(Of ButtonGroupBuilder)
		Implements IVisible(Of ButtonGroupBuilder)

		Private ReadOnly _builder As ControlBuilder
		Private ReadOnly _controls As ISet(Of RibbonElement)

		Public Sub New()
			Dim defaultProvider As IDefaultProvider = New Defaults(Of ButtonGroup)
			Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
			attributeGroupBuilder.SetDefaults(defaultProvider)
			_builder = New ControlBuilder(attributeGroupBuilder)
			_controls = New HashSet(Of RibbonElement)
		End Sub

		Public Function Build(Optional tag As Object = Nothing) As ButtonGroup
			Return New ButtonGroup(_controls, _builder.Build(), tag:=tag)
		End Function

		Public Function WithControls(ParamArray controls() As Button) As ButtonGroupBuilder
			Array.ForEach(controls, Sub(c) _controls.Add(c))
			Return Me
		End Function

		Public Function WithControls(ParamArray controls() As Gallery) As ButtonGroupBuilder
			Array.ForEach(controls, Sub(c) _controls.Add(c))
			Return Me
		End Function

		Public Function WithControls(ParamArray controls() As Menu) As ButtonGroupBuilder
			Array.ForEach(controls, Sub(c) _controls.Add(c))
			Return Me
		End Function

		Public Function WithControls(ParamArray controls() As Separator) As ButtonGroupBuilder
			Array.ForEach(controls, Sub(c) _controls.Add(c))
			Return Me
		End Function

		Public Function WithControls(ParamArray controls() As SplitButton) As ButtonGroupBuilder
			Array.ForEach(controls, Sub(c) _controls.Add(c))
			Return Me
		End Function

		Public Function WithControls(ParamArray controls() As ToggleButton) As ButtonGroupBuilder
			Array.ForEach(controls, Sub(c) _controls.Add(c))
			Return Me
		End Function

		Public Function Visible() As ButtonGroupBuilder Implements IVisible(Of ButtonGroupBuilder).Visible
			_builder.Visible()
			Return Me
		End Function

		Public Function Visible(callback As FromControl(Of Boolean)) As ButtonGroupBuilder Implements IVisible(Of ButtonGroupBuilder).Visible
			_builder.Visible(callback)
			Return Me
		End Function

		Public Function Invisible() As ButtonGroupBuilder Implements IVisible(Of ButtonGroupBuilder).Invisible
			_builder.Invisible()
			Return Me
		End Function

		Public Function Invisible(callback As FromControl(Of Boolean)) As ButtonGroupBuilder Implements IVisible(Of ButtonGroupBuilder).Invisible
			_builder.Invisible(callback)
			Return Me
		End Function

		Public Function InsertBeforeMSO(builtInControl As MSO) As ButtonGroupBuilder Implements IInsert(Of ButtonGroupBuilder).InsertBefore
			_builder.InsertBefore(builtInControl)
			Return Me
		End Function

		Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As ButtonGroupBuilder Implements IInsert(Of ButtonGroupBuilder).InsertBefore
			_builder.InsertBefore(qualifiedControl)
			Return Me
		End Function

		Public Function InsertAfterMSO(builtInControl As MSO) As ButtonGroupBuilder Implements IInsert(Of ButtonGroupBuilder).InsertAfter
			_builder.InsertAfter(builtInControl)
			Return Me
		End Function

		Public Function InsertAfterQ(qualifiedControl As RibbonElement) As ButtonGroupBuilder Implements IInsert(Of ButtonGroupBuilder).InsertAfter
			_builder.InsertAfter(qualifiedControl)
			Return Me
		End Function
	End Class
End Namespace