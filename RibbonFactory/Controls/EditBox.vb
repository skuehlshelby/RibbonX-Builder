Imports RibbonFactory.Component_Interfaces
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
        Implements IOnChange

        Private ReadOnly _attributes As AttributeGroup

        Friend Sub New(attributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(tag)
            _attributes = attributes
            AddHandler _attributes.AttributeChanged, AddressOf OnValueChanged
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
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.Enabled).SetValue(value)
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Visible).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.Visible).SetValue(value)
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Label).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeName.Label).SetValue(value)
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.Lookup(Of Screentip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetScreentip).SetValue(Value)
                
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return _attributes.Lookup (Of SuperTip).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetSuperTip).SetValue(value)
                
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return _attributes.Lookup (Of ShowLabel).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetShowLabel).SetValue(value)
                
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
                Return _attributes.Lookup(Of GetImage).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetImage).SetValue(value)
                
            End Set
        End Property

        Public Property ShowImage As Boolean Implements IShowImage.ShowImage
            Get
                Return _attributes.Lookup (Of ShowImage).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetShowImage).SetValue(value)
                
            End Set
        End Property

        Public Property Text As String Implements IText.Text
            Get
                Return _attributes.Lookup (Of GetText).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetText).SetValue(value)
                
            End Set
        End Property

        Public ReadOnly Property MaxLength As UShort Implements IMaxLength.MaxLength
            Get
                Return _attributes.Lookup (Of MaxLength).GetValue()
            End Get
        End Property

        Public Sub Execute() Implements IOnChange.Execute
            _attributes.Lookup (Of OnChange).GetValue().Invoke()
        End Sub
        
    End Class

End Namespace