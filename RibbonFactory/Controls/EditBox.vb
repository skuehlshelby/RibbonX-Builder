Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.Utilities.Validation
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
        Implements IOnAction

        Private ReadOnly _attributes As AttributeSet
        Private ReadOnly _validationRules As ICollection(Of IValidate(Of String))
        Private ReadOnly _onTextChange As ICollection(Of Action(Of String))
        Private ReadOnly _onValidationError As ICollection(Of Action(Of String))

        Public Event TextChanged As EventHandler(Of TextChangedEventArgs)

        Friend Sub New(attributes As AttributeSet, validationRules As ICollection(Of IValidate(Of String)), Optional tag As Object = Nothing)
            MyBase.New(tag)
            _attributes = attributes
            _validationRules = validationRules
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
                Try
                    SuspendAutomaticRefreshing()

                    For Each rule As IValidate(Of String) In _validationRules
                        Dim validationResult As IValidationResult = rule.Validate(Value)

                        If Not validationResult.Success Then
                            RaiseEvent TextChanged(Me, TextChangedEventArgs.Failure(validationResult.FailureMessage))
                            Exit Property
                        End If
                    Next

                    _attributes.ReadWriteLookup(Of String)(AttributeCategory.Text).SetValue(Value)
                    RaiseEvent TextChanged(Me, TextChangedEventArgs.Success(Value))
                Finally
                    ResumeAutomaticRefreshing()
                End Try
            End Set
        End Property

        Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Return _attributes
        End Function

        Public Sub Execute() Implements IOnAction.Execute
            _attributes.ReadOnlyLookup(Of Action(Of EditBox))(AttributeCategory.OnChange).GetValue().Invoke(Me)
        End Sub

        Public ReadOnly Property MaxLength As Integer Implements IMaxLength.MaxLength
            Get
                Return _attributes.ReadOnlyLookup(Of Integer)(AttributeName.MaxLength).GetValue()
            End Get
        End Property

        Public Class TextChangedEventArgs
            Inherits EventArgs

            Private Sub New(text As String, errorMessage As String)
                Me.ErrorMessage = errorMessage
                Me.Text = text
            End Sub

            Public ReadOnly Property IsSuccess As Boolean
                Get
                    Return ErrorMessage.Equals(String.Empty)
                End Get
            End Property

            Public ReadOnly Property IsFailure As Boolean
                Get
                    Return Not IsSuccess
                End Get
            End Property

            Public ReadOnly Property ErrorMessage As String

            Public ReadOnly Property Text As String

            Public Shared Function Success(text As String) As TextChangedEventArgs
                Return New TextChangedEventArgs(text, String.Empty)
            End Function

            Public Shared Function Failure(errorMessage As String) As TextChangedEventArgs
                Return New TextChangedEventArgs(String.Empty, errorMessage)
            End Function

        End Class

    End Class

End Namespace