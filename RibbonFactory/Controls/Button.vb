Public Class Button
    Inherits RibbonElement
    Implements IActionable
    Implements IDisablable
    Implements IDescribable
    Implements IVeryDescribable
    Implements IGraphic
    Implements IResizable

    Private ReadOnly ButtonAttributes As ControlPropertyList

    Friend Sub New(ButtonAttributes As ControlPropertyList, Optional ByVal Tag As Object = Nothing)
        MyBase.New(RibbonElement.FabricateID(Of Button)(), Tag)
        Me.ButtonAttributes = ButtonAttributes
    End Sub

    Public Overrides ReadOnly Property XML As String
        Get
            Return String.Format(String.Join(vbNewLine, "<button id=""{0}""", "{1}", "></button>"), ID, String.Join(vbNewLine & vbTab, ButtonAttributes.Select(Function(A) A.XML)))
        End Get
    End Property

    Public Property Description As String Implements IVeryDescribable.Description
        Get
            Return ButtonAttributes.GetValue(Of String, Attributes.Description)
        End Get
        Set(value As String)
            ButtonAttributes.SetValue(Of String, Attributes.Description)(value)
        End Set
    End Property

    Public Property Label As String Implements IDescribable.Label
        Get
            Return ButtonAttributes.GetValue(Of String, Attributes.Label)
        End Get
        Set(value As String)
            ButtonAttributes.SetValue(Of String, Attributes.Label)(value)
        End Set
    End Property

    Public Property ShowLabel As Boolean Implements IDescribable.ShowLabel
        Get
            Return ButtonAttributes.GetValue(Of Boolean, Attributes.ShowLabel)
        End Get
        Set(value As Boolean)
            ButtonAttributes.SetValue(Of Boolean, Attributes.ShowLabel)(value)
        End Set
    End Property

    Public Property ScreenTip As String Implements IDescribable.ScreenTip
        Get
            Return ButtonAttributes.GetValue(Of String, Attributes.Screentip)
        End Get
        Set(value As String)
            ButtonAttributes.SetValue(Of String, Attributes.Screentip)(value)
        End Set
    End Property

    Public Property Supertip As String Implements IDescribable.Supertip
        Get
            Return ButtonAttributes.GetValue(Of String, Attributes.Supertip)
        End Get
        Set(value As String)
            ButtonAttributes.SetValue(Of String, Attributes.Supertip)(value)
        End Set
    End Property

    Public Property Enabled As Boolean Implements IDisablable.Enabled
        Get
            Return ButtonAttributes.GetValue(Of Boolean, Attributes.Enabled)
        End Get
        Set(value As Boolean)
            ButtonAttributes.SetValue(Of Boolean, Attributes.Enabled)(value)
        End Set
    End Property

    Public Property Visible As Boolean Implements IDisablable.Visible
        Get
            Return ButtonAttributes.GetValue(Of Boolean, Attributes.Visible)
        End Get
        Set(value As Boolean)
            ButtonAttributes.SetValue(Of Boolean, Attributes.Visible)(value)
        End Set
    End Property

    Public Property Image As Object Implements IGraphic.Image
        Get
            Return ButtonAttributes.GetValue(Of Object, Attributes.Image)
        End Get
        Set(value As Object)
            ButtonAttributes.SetValue(Of Object, Attributes.Image)(value)
        End Set
    End Property

    Public ReadOnly Property IsCustom As Boolean Implements IGraphic.IsCustom
        Get
            Return TypeOf ButtonAttributes.GetValue(Of Object, Attributes.Image) Is stdole.IPictureDisp
        End Get
    End Property

    Public Property ShowImage As Boolean Implements IGraphic.ShowImage
        Get
            Return ButtonAttributes.GetValue(Of Boolean, Attributes.ShowImage)
        End Get
        Set(value As Boolean)
            ButtonAttributes.SetValue(Of Boolean, Attributes.ShowImage)(value)
        End Set
    End Property

    Public Property Size As ControlSize Implements IResizable.Size
        Get
            Return ButtonAttributes.GetValue(Of ControlSize, Attributes.Size)
        End Get
        Set(value As ControlSize)
            ButtonAttributes.SetValue(Of ControlSize, Attributes.Size)(value)
        End Set
    End Property

    Public Sub Execute() Implements IActionable.Execute
        ButtonAttributes.GetValue(Of Action, Attributes.Action).Invoke()
    End Sub

    Public Overrides Function Equals(obj As Object) As Boolean
        If obj.GetHashCode() = GetHashCode() Then
            Return TypeOf obj Is Button
        Else
            Return False
        End If
    End Function
End Class