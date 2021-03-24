Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.Enabled
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

        Private ReadOnly Attributes As AttributeGroup

        Friend Sub New(buttonAttributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(tag)
        End Sub

        Public Overrides ReadOnly Property ID As String

        Public Overrides ReadOnly Property XML As String
            Get
                Return _
                    String.Format(String.Join(vbNewLine, "<labelControl id=""{0}""", "{1}", "></labelControl>"), ID,
                                  String.Join(vbNewLine & vbTab, Attributes.Select(Function(a) a.XML)))
            End Get
        End Property

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return Attributes.Lookup(Of Enabled).GetValue()
            End Get
            Set
                Attributes.Lookup(Of GetEnabled).SetValue(value)
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return Attributes.Lookup(Of Visible).GetValue()
            End Get
            Set
                Attributes.Lookup(Of GetVisible).SetValue(value)
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return Attributes.Lookup(Of Label).GetValue()
            End Get
            Set
                Attributes.Lookup(Of GetLabel).SetValue(value)
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return Attributes.Lookup(Of ShowLabel).GetValue()
            End Get
            Set
                Attributes.Lookup(Of GetShowLabel).SetValue(value)
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return Attributes.Lookup(Of Screentip).GetValue()
            End Get
            Set
                Attributes.Lookup(Of GetScreentip).SetValue(value)
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.Supertip
            Get
                Return Attributes.Lookup(Of Supertip).GetValue()
            End Get
            Set
                Attributes.Lookup(Of GetSupertip).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Overrides Function Equals(obj As Object) As Boolean
            Return obj.GetHashCode() = GetHashCode() AndAlso TypeOf obj Is Button
        End Function
    End Class
End NameSpace