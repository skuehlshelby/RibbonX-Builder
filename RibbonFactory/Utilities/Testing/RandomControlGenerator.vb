Imports System.Reflection
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Builders
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.Utilities

Public NotInheritable Class RandomControlGenerator

	Private ReadOnly random As Random = New Random()
	Private ReadOnly callbacks As ICreateCallbacks

	Public Sub New(callbacks As ICreateCallbacks)
		Me.callbacks = callbacks
	End Sub

	Public Function MakeButton() As Button
		Return MakeControl(Function() New ButtonBuilder())
	End Function

	Public Function MakeCheckbox() As CheckBox
		Return MakeControl(Function() New CheckBoxBuilder())
	End Function

	Public Function MakeComboBox() As ComboBox
		Return MakeControl(Function() New ComboBoxBuilder())
	End Function

	Public Function MakeEditBox() As EditBox
		Return MakeControl(Function() New EditBoxBuilder())
	End Function

	Public Function MakeLabelControl() As LabelControl
		Return MakeControl(Function() New LabelControlBuilder())
	End Function

	Public Function MakeToggleButton() As ToggleButton
		Return MakeControl(Function() New ToggleButtonBuilder())
	End Function

	Public Function MakeSplitButtonWithRegularButtonAndMenu() As SplitButton
		Dim builder As SplitButtonBuilder = New SplitButtonBuilder()

		ConfigureControl(builder)

		builder.WithButton(MakeButton())

		Dim menuBuilder As MenuBuilder = New MenuBuilder()

		ConfigureControl(menuBuilder)

		menuBuilder.WithControls(MakeButton())

		builder.WithMenu(menuBuilder.Build())

		Return builder.Build()
	End Function

	Public Function MakeDropDown() As DropDown
		Return MakeControl(Function() New DropDownBuilder())
	End Function

	Public Function MakeGallery() As Gallery
		Return MakeControl(Function() New GalleryBuilder())
	End Function

	Public Function MakeMenu() As Menu
		Return MakeControl(Function() New MenuBuilder())
	End Function

	Private Function MakeControl(Of T As RibbonElement)(builderInitializer As Func(Of IBuilder(Of T))) As T
		Dim builder As IBuilder(Of T) = builderInitializer.Invoke()

		ConfigureControl(builder)

		Return builder.Build()
	End Function

	Private Sub ConfigureControl(Of T As RibbonElement)(builder As IBuilder(Of T))
		Dim type As Type = builder.GetType()

		Dim method As MethodInfo = Nothing

		If builder.GetType().ImplementsGenericInterface(GetType(ILabelOnly(Of ))) Then
			method = type.GetMethod(NameOf(ILabelOnly(Of Object).WithLabel), New Type() {})

			method.Invoke(builder, New Object() {"A Sweet Label"})
		End If

		If builder.GetType().ImplementsGenericInterface(GetType(ILabelScreenTipSuperTip(Of ))) Then
			If random.NextBoolean() Then
				method = type.GetMethod(NameOf(ILabelScreenTipSuperTip(Of Object).WithLabel), New Type() {GetType(String), GetType(Boolean)})

				method.Invoke(builder, New Object() {"A Sweet Label", random.NextBoolean()})
			Else
				method = type.GetMethod(NameOf(ILabelScreenTipSuperTip(Of Object).WithLabel), New Type() {GetType(String), GetType(FromControl(Of String)), GetType(Boolean)})

				Dim callback As FromControl(Of String) = New FromControl(Of String)(AddressOf callbacks.GetLabel)

				method.Invoke(builder, New Object() {"A Different, But Still Sweet Label", callback, random.NextBoolean()})
			End If

			If random.NextBoolean() Then
				If random.NextBoolean() Then
					method = type.GetMethod(NameOf(ILabelScreenTipSuperTip(Of Object).WithScreenTip), New Type() {GetType(String)})

					method.Invoke(builder, New Object() {"The Screentip"})
				Else
					method = type.GetMethod(NameOf(ILabelScreenTipSuperTip(Of Object).WithScreenTip), New Type() {GetType(String), GetType(FromControl(Of String))})

					Dim callback As FromControl(Of String) = New FromControl(Of String)(AddressOf callbacks.GetScreenTip)

					method.Invoke(builder, New Object() {"A Different Screentip", callback})
				End If
			End If

			If random.NextBoolean() Then
				method = type.GetMethod(NameOf(ILabelScreenTipSuperTip(Of Object).WithSuperTip), New Type() {GetType(String)})

				method.Invoke(builder, New Object() {"The Supertip"})
			Else
				method = type.GetMethod(NameOf(ILabelScreenTipSuperTip(Of Object).WithSuperTip), New Type() {GetType(String), GetType(FromControl(Of String))})

				Dim callback As FromControl(Of String) = New FromControl(Of String)(AddressOf callbacks.GetSuperTip)

				method.Invoke(builder, New Object() {"A Different Supertip", callback})
			End If
		End If

		If builder.GetType().ImplementsGenericInterface(GetType(IDescription(Of ))) Then
			If random.NextBoolean() Then
				If random.NextBoolean() Then
					method = type.GetMethod(NameOf(IDescription(Of Object).WithDescription), New Type() {GetType(String)})

					method.Invoke(builder, New Object() {"The Description"})
				Else
					method = type.GetMethod(NameOf(IDescription(Of Object).WithDescription), New Type() {GetType(String), GetType(FromControl(Of String))})

					Dim callback As FromControl(Of String) = New FromControl(Of String)(AddressOf callbacks.GetDescription)

					method.Invoke(builder, New Object() {"A Different Description", callback})
				End If
			End If
		End If

		If builder.GetType().ImplementsGenericInterface(GetType(IKeyTip(Of ))) Then
			If random.NextBoolean() Then
				If random.NextBoolean() Then
					method = type.GetMethod(NameOf(IKeyTip(Of Object).WithKeyTip), New Type() {GetType(KeyTip)})

					method.Invoke(builder, New Object() {KeyTip.NextAvailable()})
				Else
					method = type.GetMethod(NameOf(IKeyTip(Of Object).WithKeyTip), New Type() {GetType(KeyTip), GetType(FromControl(Of KeyTip))})

					Dim callback As FromControl(Of KeyTip) = New FromControl(Of KeyTip)(AddressOf callbacks.GetKeyTip)

					method.Invoke(builder, New Object() {KeyTip.NextAvailable(), callback})
				End If
			End If
		End If

		If builder.GetType().ImplementsGenericInterface(GetType(IVisible(Of ))) Then
			If random.NextBoolean() Then
				If random.NextBoolean() Then
					If random.NextBoolean() Then
						method = type.GetMethod(NameOf(IVisible(Of Object).Visible), New Type() {})

						method.Invoke(builder, New Object() {})
					Else
						method = type.GetMethod(NameOf(IVisible(Of Object).Visible), New Type() {GetType(FromControl(Of Boolean))})

						Dim callback As FromControl(Of Boolean) = New FromControl(Of Boolean)(AddressOf callbacks.GetVisible)

						method.Invoke(builder, New Object() {callback})
					End If
				Else
					If random.NextBoolean() Then
						method = type.GetMethod(NameOf(IVisible(Of Object).Invisible), New Type() {})

						method.Invoke(builder, New Object() {})
					Else
						method = type.GetMethod(NameOf(IVisible(Of Object).Invisible), New Type() {GetType(FromControl(Of Boolean))})

						Dim callback As FromControl(Of Boolean) = New FromControl(Of Boolean)(AddressOf callbacks.GetVisible)

						method.Invoke(builder, New Object() {callback})
					End If
				End If
			End If
		End If

		If builder.GetType().IsDerivedFromGenericType(GetType(IEnabled(Of ))) Then
			If random.NextBoolean() Then
				If random.NextBoolean() Then
					If random.NextBoolean() Then
						method = type.GetMethod(NameOf(IEnabled(Of Object).Enabled), New Type() {})

						method.Invoke(builder, New Object() {})
					Else
						method = type.GetMethod(NameOf(IEnabled(Of Object).Enabled), New Type() {GetType(FromControl(Of Boolean))})

						Dim callback As FromControl(Of Boolean) = New FromControl(Of Boolean)(AddressOf callbacks.GetEnabled)

						method.Invoke(builder, New Object() {callback})
					End If
				Else
					If random.NextBoolean() Then
						method = type.GetMethod(NameOf(IEnabled(Of Object).Disabled), New Type() {})

						method.Invoke(builder, New Object() {})
					Else
						method = type.GetMethod(NameOf(IEnabled(Of Object).Disabled), New Type() {GetType(FromControl(Of Boolean))})

						Dim callback As FromControl(Of Boolean) = New FromControl(Of Boolean)(AddressOf callbacks.GetEnabled)

						method.Invoke(builder, New Object() {callback})
					End If
				End If
			End If
		End If

		If builder.GetType().ImplementsGenericInterface(GetType(ISize(Of ))) Then
			If random.NextBoolean() Then
				If random.NextBoolean() Then
					If random.NextBoolean() Then
						method = type.GetMethod(NameOf(ISize(Of Object).Normal), New Type() {})

						method.Invoke(builder, New Object() {})
					Else
						method = type.GetMethod(NameOf(ISize(Of Object).Normal), New Type() {GetType(FromControl(Of ControlSize))})

						Dim callback As FromControl(Of ControlSize) = New FromControl(Of ControlSize)(AddressOf callbacks.GetSize)

						method.Invoke(builder, New Object() {callback})
					End If
				Else
					If random.NextBoolean() Then
						method = type.GetMethod(NameOf(ISize(Of Object).Large), New Type() {})

						method.Invoke(builder, New Object() {})
					Else
						method = type.GetMethod(NameOf(ISize(Of Object).Large), New Type() {GetType(FromControl(Of ControlSize))})

						Dim callback As FromControl(Of ControlSize) = New FromControl(Of ControlSize)(AddressOf callbacks.GetSize)

						method.Invoke(builder, New Object() {callback})
					End If
				End If
			End If
		End If

		If builder.GetType().ImplementsGenericInterface(GetType(IMaxLength(Of ))) Then
			If random.NextBoolean() Then
				method = type.GetMethod(NameOf(IMaxLength(Of Object).WithMaximumInputLength), New Type() {GetType(Byte)})

				method.Invoke(builder, New Object() {CByte(random.Next(0, 50))})
			End If
		End If

		If builder.GetType().ImplementsGenericInterface(GetType(ISizeString(Of ))) Then
			If random.NextBoolean() Then
				method = type.GetMethod(NameOf(ISizeString(Of Object).AsWideAs), New Type() {GetType(String)})

				method.Invoke(builder, New Object() {New String("a"c, random.Next(0, 50))})
			End If
		End If

		If builder.GetType().ImplementsGenericInterface(GetType(IShowLabel(Of ))) Then
			If random.NextBoolean() Then
				If random.NextBoolean() Then
					method = type.GetMethod(NameOf(IShowLabel(Of Object).ShowLabel), New Type() {})

					method.Invoke(builder, New Object() {})
				Else
					method = type.GetMethod(NameOf(IShowLabel(Of Object).HideLabel), New Type() {})

					method.Invoke(builder, New Object() {})
				End If
			End If
		End If

	End Sub

End Class
