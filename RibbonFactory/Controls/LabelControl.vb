Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.Enabled
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.Screentip
Imports RibbonFactory.RibbonAttributes.Categories.Supertip
Imports RibbonFactory.RibbonAttributes.Categories.Visible

Namespace Controls
    Public Class LabelControl
        Inherits RibbonElement
        Implements IEnabled
        Implements IVisible
        Implements ILabel
        Implements IShowLabel
        Implements IScreenTip
        Implements ISuperTip

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
                Return $"<checkBox { String.Join(" ", _attributes) }/>"
            End Get
        End Property

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return _attributes.Lookup(Of Enabled).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetEnabled).SetValue(value)
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.Lookup(Of Visible).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetVisible).SetValue(value)
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.Lookup(Of Label).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetLabel).SetValue(value)
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return _attributes.Lookup(Of ShowLabel).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetShowLabel).SetValue(value)
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.Lookup(Of Screentip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetScreentip).SetValue(value)
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.Supertip
            Get
                Return _attributes.Lookup(Of Supertip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetSupertip).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Overrides Function Equals(obj As Object) As Boolean
            Return obj.GetHashCode() = GetHashCode() AndAlso TypeOf obj Is Button
        End Function
    End Class
End NameSpace