Friend MustInherit Class ControlProperty
    Implements IControlProperty
    Protected Const XML_TEMPLATE As String = "{0}=""{1}"""

    Private Structure TThis
        Public Governs As Attributes.AttributeGroup
    End Structure

    Private This As TThis

    Protected Sub New(ByVal Governs As Attributes.AttributeGroup)
        This = New TThis With {
                .Governs = Governs
        }
    End Sub
    Friend ReadOnly Property Governs As Attributes.AttributeGroup Implements IControlProperty.Governs
        Get
            Return This.Governs
        End Get
    End Property

    Friend MustOverride ReadOnly Property XML As String Implements IControlProperty.XML

    Public Overrides Function ToString() As String Implements IControlProperty.ToString
        Return XML
    End Function
    Public Overrides Function Equals(obj As Object) As Boolean Implements IControlProperty.Equals
        If obj.GetHashCode = Me.GetHashCode Then
            Return obj IsNot Nothing AndAlso TypeOf obj Is ControlProperty AndAlso DirectCast(obj, ControlProperty).Governs.Equals(Me.Governs)
        Else
            Return False
        End If
    End Function
    Public Overrides Function GetHashCode() As Integer Implements IControlProperty.GetHashCode
        Return This.Governs.GetHashCode()
    End Function
End Class
Friend Class ControlProperty(Of T)
    Inherits ControlProperty
    Implements IControlProperty(Of T)

    Private Structure TThis
        Public Value As T
    End Structure

    Private This As TThis
    Friend Sub New(ByVal Governs As Attributes.AttributeGroup, ByVal Value As T)
        MyBase.New(Governs)

        This = New TThis With {
            .Value = Value
        }
    End Sub
    Friend Overrides ReadOnly Property XML As String
        Get
            Return String.Format(XML_TEMPLATE, Governs, This.Value.ToString())
        End Get
    End Property

    Public Function GetValue() As T Implements IControlProperty(Of T).GetValue
        Return This.Value
    End Function
End Class

Friend Class DynamicControlProperty(Of T)
    Inherits ControlProperty
    Implements IDynamicControlProperty(Of T)
    Private Structure TThis
        Public Value As T
        Public CallbackName As String
    End Structure

    Private This As TThis
    Friend Sub New(ByVal Governs As Attributes.AttributeGroup, ByVal CallbackName As String, ByVal InitialValue As T)
        MyBase.New(Governs)

        This = New TThis With {
            .CallbackName = CallbackName,
            .Value = InitialValue
        }
    End Sub
    Friend Overrides ReadOnly Property XML As String
        Get
            Return String.Format(XML_TEMPLATE, Governs, This.CallbackName)
        End Get
    End Property
    Public Function GetValue() As T Implements IControlProperty(Of T).GetValue
        Return This.Value
    End Function
    Public Sub SetValue(Value As T) Implements IDynamicControlProperty(Of T).SetValue
        If Not This.Value.Equals(Value) Then
            This.Value = Value
        End If
    End Sub
End Class