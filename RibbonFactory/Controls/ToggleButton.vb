Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.Description
Imports RibbonFactory.RibbonAttributes.Categories.Enabled
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.Image
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.Pressed
Imports RibbonFactory.RibbonAttributes.Categories.Screentip
Imports RibbonFactory.RibbonAttributes.Categories.Size
Imports RibbonFactory.RibbonAttributes.Categories.Supertip
Imports RibbonFactory.RibbonAttributes.Categories.Visible

Namespace Controls

    Public Class ToggleButton
        Inherits RibbonElement
        Implements IOnAction
        Implements IEnabled
        Implements IScreenTip
        Implements ISuperTip
        Implements IDescription
        Implements ILabel
        Implements IShowLabel
        Implements IImage
        Implements IShowImage
        Implements ISize
        Implements IPressed
        Implements IVisible

        Private ReadOnly _attributes As AttributeGroup

        Friend Sub New(buttonAttributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(tag)
            _attributes = buttonAttributes
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.Lookup(Of Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return $"<toggleButton { String.Join(" ", _attributes) }/>"
            End Get
        End Property

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return _attributes.Lookup(Of Enabled).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetEnabled).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.Lookup(Of Visible).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetVisible).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.Lookup(Of Label).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetLabel).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return _attributes.Lookup(Of ShowLabel).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetShowLabel).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.Lookup(Of Screentip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetScreentip).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property SuperTip As String Implements ISupertip.Supertip
            Get
                Return _attributes.Lookup(Of Supertip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetSuperTip).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property Size As Enums.ControlSize Implements ISize.Size
            Get
                Return _attributes.Lookup (Of Size).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetSize).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property Pressed As Boolean Implements IPressed.Pressed
            Get
                Return _attributes.Lookup (Of GetPressed).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetPressed).SetValue(value)
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
                    _attributes.Lookup(Of GetImage).SetValue(CType(value, stdole.IPictureDisp))
                    Refresh()
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

        Public Property Description As String Implements IDescription.Description
            Get
                Return _attributes.Lookup(Of Description).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetDescription).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Sub Execute() Implements IOnAction.Execute
            _attributes.Lookup(Of Categories.OnAction.OnAction).GetValue().Invoke()
        End Sub

        Public Overrides Function ToString() As String
            Return XML
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return obj.GetHashCode() = GetHashCode() AndAlso TypeOf obj Is ToggleButton
        End Function
    End Class
End NameSpace