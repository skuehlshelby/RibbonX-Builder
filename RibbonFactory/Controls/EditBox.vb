Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Controls

    Public Class EditBox
        Inherits RibbonElement
        Implements IEnabled
        Implements IVisible
        Implements ILabel
        Implements IScreenTip
        Implements ISuperTip
        Implements IShowLabel
        Implements IKeyTip
        Implements IImage
        Implements IShowImage
        Implements IText
        Implements IMaxLength
        Implements IDefaultProvider

        Private ReadOnly _attributes As AttributeSet

        Public Event BeforeTextChange As EventHandler(Of BeforeTextChangeEventArgs)

        Public Event TextChanged As EventHandler(Of TextChangedEventArgs)

        Friend Sub New(attributes As AttributeSet, Optional tag As Object = Nothing)
            MyBase.New(tag)
            _attributes = attributes
            AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return $"<editBox { _attributes }/>"
            End Get
        End Property

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Enabled).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.Enabled).SetValue(Value)
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Visible).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.Visible).SetValue(Value)
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Label).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeName.Label).SetValue(Value)
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

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.ShowLabel).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetShowLabel).SetValue(Value)
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

        Public Property Image As IPictureDisp Implements IImage.Image
            Get
                Return _attributes.ReadOnlyLookup(Of IPictureDisp)(AttributeCategory.Image).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of IPictureDisp)(AttributeCategory.Image).SetValue(Value)
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

        Public Property Text As String Implements IText.Text
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.GetText).GetValue()
            End Get
            Set
                Dim oldValue As String = Text

                If Not oldValue.Equals(Value, StringComparison.OrdinalIgnoreCase) Then
                    Try
                        SuspendAutomaticRefreshing()

                        Dim e As BeforeTextChangeEventArgs = New BeforeTextChangeEventArgs(Value)

                        RaiseEvent BeforeTextChange(Me, e)

                        If Not e.Cancel Then
                            _attributes.ReadWriteLookup(Of String)(AttributeCategory.Text).SetValue(e.NewText)

                            RaiseEvent TextChanged(Me, New TextChangedEventArgs(Value, oldValue))
                        End If
                    Finally
                        ResumeAutomaticRefreshing()
                    End Try
                End If
            End Set
        End Property

        Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Return _attributes
        End Function

        Public ReadOnly Property MaxLength As Integer Implements IMaxLength.MaxLength
            Get
                Return _attributes.ReadOnlyLookup(Of Integer)(AttributeName.MaxLength).GetValue()
            End Get
        End Property

        Public Class TextChangedEventArgs
            Inherits EventArgs

            Public Sub New(newText As String, oldText As String)
                Me.NewText = newText
                Me.OldText = oldText
            End Sub

            Public Property NewText As String

            Public ReadOnly Property OldText As String

        End Class

        Public NotInheritable Class BeforeTextChangeEventArgs
            Inherits EventArgs

            Public Sub New(newText As String)
                Me.NewText = newText
                Cancel = False
            End Sub

            Public ReadOnly Property NewText As String

            Public Property Cancel As Boolean

        End Class

    End Class

End Namespace