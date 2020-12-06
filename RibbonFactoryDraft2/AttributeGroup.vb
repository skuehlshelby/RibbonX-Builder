Public Class AttributeGroup(Of T)
    Friend ReadOnly Property Name As String
    Friend ReadOnly Property Type As T
    Friend ReadOnly Property Required As Boolean
End Class

Friend Class Label
    Public ReadOnly Property Label As AttributeGroup(Of STString) = New AttributeGroup(Of STString)()
End Class
Friend Class RibbonXLine
    Protected Sub New(ByVal Name As String, ByVal Required As Boolean)
        Me.Name = Name
        Me.Required = Required
    End Sub
    Friend ReadOnly Property Name As String
    Friend ReadOnly Property Required As Boolean
End Class

Friend Class RibbonXLine(Of T)
    Inherits RibbonXLine
    Friend Sub New(ByVal Name As String, Value As T, Required As Boolean)
        MyBase.New(Name, Required)
        Me.Value = Value
    End Sub

    Friend ReadOnly Property Value As T
End Class

Friend Class RestrictedLengthString
    Friend Sub New(ByVal Value As String, MinLength As Byte, MaxLength As UShort)
        If Value.Length < MinLength OrElse Value.Length > MaxLength Then
            If Value.Length < MinLength Then
                Throw New ArgumentException(String.Format("{0} is less than the allowable length of {1} character{2}.", Value, MinLength, If(MinLength > 1, "s", String.Empty)))
            Else
                Throw New ArgumentException(String.Format("{0} exceeds the allowable length of {1} character{2}.", Value, MaxLength, If(MaxLength > 1, "s", String.Empty)))
            End If
        Else
            Me.Value = Value
        End If
    End Sub
    Friend ReadOnly Property Value As String
End Class

Friend Class STString
    Inherits RestrictedLengthString
    Private Const MIN_LENGTH As Byte = 1
    Private Const MAX_LENGTH As UShort = 1024
    Friend Sub New(ByVal Value As String)
        MyBase.New(Value, MIN_LENGTH, MAX_LENGTH)
    End Sub
End Class
Friend Class STDelegate
    Inherits STString
    Friend Sub New(ByVal Value As String)
        MyBase.New(Value)
    End Sub
End Class
Friend Class STID
    Inherits STString
    Friend Sub New(ByVal Value As String)
        MyBase.New(Value)
    End Sub
End Class
Friend Class STLongString
    Inherits RestrictedLengthString
    Private Const MIN_LENGTH As Byte = 1
    Private Const MAX_LENGTH As UShort = 4096
    Friend Sub New(ByVal Value As String)
        MyBase.New(Value, MIN_LENGTH, MAX_LENGTH)
    End Sub
End Class

Friend Class STBoxStyle
    Friend Enum BoxStyle As Byte
        horizontal = 0
        vertical = 1
    End Enum

    Friend Sub New(ByVal Style As BoxStyle)
        Me.Value = Style
    End Sub
    Friend ReadOnly Property Value As BoxStyle
End Class
