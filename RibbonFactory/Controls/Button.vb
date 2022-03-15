Imports RibbonFactory.Builders
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Controls

	Public NotInheritable Class Button
		Inherits RibbonElement
		Implements IEnabled
		Implements IVisible
		Implements ILabel
		Implements IShowLabel
		Implements IScreenTip
		Implements ISuperTip
		Implements IKeyTip
		Implements IDescription
		Implements IImage
		Implements IOnAction
		Implements IShowImage
		Implements ISize
		Implements IDefaultProvider

		Private ReadOnly _attributes As AttributeSet

		Public Sub New(steps As Action(Of IButtonBuilder), Optional tag As Object = Nothing)
			MyBase.New(tag)

			Dim builder As ButtonBuilder = New ButtonBuilder()

			steps.Invoke(builder)

			_attributes = builder.Build()
		End Sub

		Public Sub New(template As RibbonElement, steps As Action(Of IButtonBuilder), Optional tag As Object = Nothing)
			MyBase.New(tag)

			Dim builder As ButtonBuilder = New ButtonBuilder(template)

			steps.Invoke(builder)

			_attributes = builder.Build()
		End Sub

#Region "Events"

		Public Custom Event BeforeClick As EventHandler(Of BeforeClickEventArgs)
		    AddHandler(value As EventHandler(Of BeforeClickEventArgs))
				_attributes.Read(Of ICollection(Of EventHandler(Of BeforeClickEventArgs)))().Add(value)
		    End AddHandler

		    RemoveHandler(value As EventHandler(Of BeforeClickEventArgs))
				_attributes.Read(Of ICollection(Of EventHandler(Of BeforeClickEventArgs)))().Remove(value)
		    End RemoveHandler

		    RaiseEvent(sender As Object, e As BeforeClickEventArgs)
				For Each target As EventHandler(Of BeforeClickEventArgs) In _attributes.Read(Of ICollection(Of EventHandler(Of BeforeClickEventArgs)))()
					Try
						target.Invoke(sender, e)
					Catch ex As Exception
						Debug.WriteLine($"Encountered an exception while invoking '{target.Method.Name}' inside '{NameOf(BeforeClick)}':")
						Debug.WriteLine(ex.Message)
					End Try
				Next
		    End RaiseEvent
		End Event

		Public Custom Event OnClick As EventHandler
		    AddHandler(value As EventHandler)
				_attributes.Read(Of ICollection(Of EventHandler))().Add(value)
		    End AddHandler

		    RemoveHandler(value As EventHandler)
				_attributes.Read(Of ICollection(Of EventHandler))().Remove(value)
		    End RemoveHandler

		    RaiseEvent(sender As Object, e As EventArgs)
				For Each target As EventHandler In _attributes.Read(Of ICollection(Of EventHandler))()
					Try
						target.Invoke(sender, e)
					Catch ex As Exception
						Debug.WriteLine($"Encountered an exception while invoking '{target.Method.Name}' inside '{NameOf(BeforeClick)}':")
						Debug.WriteLine(ex.Message)
					End Try
				Next
		    End RaiseEvent
		End Event

#End Region

#Region "Base Class Overrides"

		Public Overrides ReadOnly Property ID As String
			Get
				Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
			End Get
		End Property

		Public Overrides ReadOnly Property XML As String
			Get
				Return $"<button { _attributes }/>"
			End Get
		End Property

#End Region

#Region "Enabled and Visible"

		Public Property Enabled As Boolean Implements IEnabled.Enabled
			Get
				Return _attributes.Read(Of Boolean)(AttributeCategory.Enabled)
			End Get
			Set
				_attributes.Write(value, AttributeCategory.Enabled)
			End Set
		End Property

		Public Property Visible As Boolean Implements IVisible.Visible
			Get
				Return _attributes.Read(Of Boolean)(AttributeCategory.Visibility)
			End Get
			Set
				_attributes.Write(value, AttributeCategory.Visibility)
			End Set
		End Property

#End Region

#Region "Display Text"

		Public Property Description As String Implements IDescription.Description
			Get
				Return _attributes.Read(Of String)(AttributeCategory.Description)
			End Get
			Set
				_attributes.Write(value, AttributeCategory.Description)
			End Set
		End Property
		Public Property Label As String Implements ILabel.Label
			Get
				Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Label).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of String)(AttributeName.GetLabel).SetValue(Value)
			End Set
		End Property

		Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
			Get
				Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.ShowLabel).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetShowLabel).SetValue(Value)
			End Set
		End Property

		Public Property ScreenTip As String Implements IScreenTip.ScreenTip
			Get
				Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Screentip).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of String)(AttributeName.GetScreentip).SetValue(Value)
			End Set
		End Property
		Public Property SuperTip As String Implements ISuperTip.SuperTip
			Get
				Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Supertip).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of String)(AttributeName.GetSupertip).SetValue(Value)
			End Set
		End Property

		Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
			Get
				Return _attributes.ReadOnlyLookup(Of KeyTip)(AttributeName.Keytip).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of KeyTip)(AttributeName.GetKeytip).SetValue(Value)
			End Set
		End Property

#End Region

#Region "Image and ShowImage"

		Public Property Image As IPictureDisp Implements IImage.Image
			Get
				Return _attributes.ReadOnlyLookup(Of IPictureDisp)(AttributeName.GetImage).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of IPictureDisp)(AttributeName.GetImage).SetValue(Value)
			End Set
		End Property

		Public Property ShowImage As Boolean Implements IShowImage.ShowImage
			Get
				Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.ShowImage).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetShowImage).SetValue(Value)
			End Set
		End Property

#End Region

#Region "Size"

		Public Property Size As ControlSize Implements ISize.Size
			Get
				Return _attributes.ReadOnlyLookup(Of ControlSize)(AttributeName.Size).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of ControlSize)(AttributeName.GetSize).SetValue(Value)
			End Set
		End Property

#End Region

#Region "Action"

		Public Sub Execute() Implements IOnAction.Execute
			Dim e As BeforeClickEventArgs = New BeforeClickEventArgs()

			RaiseEvent BeforeClick(Me, e)

			If Not e.Cancel Then
				RaiseEvent OnClick(Me, EventArgs.Empty)
			End If
		End Sub

#End Region

		Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
			Return _attributes
		End Function

		Public NotInheritable Class BeforeClickEventArgs
			Inherits EventArgs

			Public Property Cancel As Boolean

		End Class

	End Class

End Namespace