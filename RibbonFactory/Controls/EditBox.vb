Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.Enabled
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.Image
Imports RibbonFactory.RibbonAttributes.Categories.KeyTip
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.MaxLength
Imports RibbonFactory.RibbonAttributes.Categories.OnChange
Imports RibbonFactory.RibbonAttributes.Categories.Screentip
Imports RibbonFactory.RibbonAttributes.Categories.SuperTip
Imports RibbonFactory.RibbonAttributes.Categories.Text
Imports RibbonFactory.RibbonAttributes.Categories.Visible
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
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.Lookup(Of Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return $"<editBox { String.Join(" ", _attributes) }/>"
            End Get
        End Property

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return _attributes.Lookup(Of Enabled).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetEnabled).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.Lookup(Of Visible).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetVisible).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.Lookup(Of Label).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetLabel).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.Lookup(Of Screentip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetScreentip).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return _attributes.Lookup (Of SuperTip).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetSuperTip).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return _attributes.Lookup (Of ShowLabel).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetShowLabel).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.Lookup (Of Categories.KeyTip.KeyTip).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetKeyTip).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property Image As Object Implements IImage.Image
            Get
                If IsCustom Then
                    Return _attributes.Lookup(Of GetImage).GetValue()
                Else
                    Return _attributes.Lookup(Of Categories.Image.ImageMso).GetValue()
                End If
            End Get
            Set
                If IsCustom Then
                    _attributes.Lookup(Of GetImage).SetValue(CType(value, IPictureDisp))
                Else
                    _attributes.Lookup(Of ImageMso).GetValue()
                End If
            End Set
        End Property

        Public ReadOnly Property IsCustom As Boolean Implements IImage.IsCustom
            Get
                Return TypeOf _attributes.Lookup(Of ImageBase) Is GetImage
            End Get
        End Property

        Public Property ShowImage As Boolean Implements IShowImage.ShowImage
            Get
                Return _attributes.Lookup (Of ShowImage).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetShowImage).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property Text As String Implements IText.Text
            Get
                Return _attributes.Lookup (Of GetText).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetText).SetValue(value)
                Refresh()
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

        Public Overrides Function Equals(obj As Object) As Boolean
            Return obj.GetHashCode() = GetHashCode() AndAlso TypeOf obj Is Button
        End Function
    End Class

End Namespace