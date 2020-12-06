Public Class ToggleButton
    Inherits RibbonElement
    Implements IActionable
    Implements IDisablable
    Implements IDescribable
    Implements IVeryDescribable
    Implements IGraphic
    Implements IResizable
    Implements IToggleable

    Private ReadOnly ButtonAttributes As ControlPropertyList

    Friend Sub New(ButtonAttributes As ControlPropertyList, Optional ByVal Tag As Object = Nothing)
        MyBase.New(FabricateID(Of ToggleButton)(), Tag)
        Me.ButtonAttributes = ButtonAttributes
    End Sub

    Public Overrides ReadOnly Property XML As String
        Get
            Return String.Format(String.Join(vbNewLine, "<toggleButton id=""{0}""", "{1}", "></button>"), ID, String.Join(vbNewLine & vbTab, ButtonAttributes.Select(Function(A) A.XML)))
        End Get
    End Property

    Public Property Enabled As Boolean Implements IDisablable.Enabled
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Boolean)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property Visible As Boolean Implements IDisablable.Visible
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Boolean)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property Label As String Implements IDescribable.Label
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As String)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property ShowLabel As Boolean Implements IDescribable.ShowLabel
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Boolean)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property ScreenTip As String Implements IDescribable.ScreenTip
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As String)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property Supertip As String Implements IDescribable.Supertip
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As String)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property Size As ControlSize Implements IResizable.Size
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As ControlSize)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property Pressed As Boolean Implements IToggleable.Pressed
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Boolean)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property Image As Object Implements IGraphic.Image
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Object)
            Throw New NotImplementedException()
        End Set
    End Property

    Public ReadOnly Property IsCustom As Boolean Implements IGraphic.IsCustom
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Property ShowImage As Boolean Implements IGraphic.ShowImage
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Boolean)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property Description As String Implements IVeryDescribable.Description
        Get
            Return ButtonAttributes.GetValue(Of String, Attributes.Description)
        End Get
        Set(value As String)
            ButtonAttributes.SetValue(Of String, Attributes.Description)(value)
        End Set
    End Property

    Public Sub Execute() Implements IActionable.Execute
        ButtonAttributes.GetValue(Of Action, Attributes.Action).Invoke()
    End Sub

    Public Overrides Function ToString() As String
        Return XML
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        If obj.GetHashCode() = GetHashCode() Then
            Return TypeOf obj Is ToggleButton
        Else
            Return False
        End If
    End Function
End Class