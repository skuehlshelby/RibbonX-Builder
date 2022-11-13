Imports System.ComponentModel
Imports RibbonX.Controls.Attributes

Namespace PrivateApi

    Friend Class RibbonProperty(Of T)
        Implements IRibbonProperty(Of T)

        Private ReadOnly callback As String
        Private value As T

        Public ReadOnly Property Name As AttributeName Implements IRibbonProperty.Name

        Public ReadOnly Property Category As AttributeCategory Implements IRibbonProperty.Category

        Public ReadOnly Property IsReadOnly As Boolean Implements IRibbonProperty.IsReadOnly

        Public ReadOnly Property IsDefault As Boolean Implements IRibbonProperty.IsDefault

        Public Function GetBoxedValue() As Object Implements IRibbonProperty.GetBoxedValue
            Return GetValue()
        End Function

        Public Function GetValue() As T Implements IRibbonProperty(Of T).GetValue
            Return value
        End Function

        Public Function SetValue(value As T) As Boolean Implements IRibbonProperty(Of T).SetValue
            If IsReadOnly Then
                Throw New InvalidOperationException()
            ElseIf Not Equals(value, Me.value) Then
                Me.value = value
                Return True
            Else
                Return False
            End If
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(TryCast(obj, IRibbonProperty))
        End Function

        Public Overloads Function Equals(other As IRibbonProperty) As Boolean Implements IEquatable(Of IRibbonProperty).Equals
            Return other IsNot Nothing AndAlso Category.Equals(other.Category)
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return Category.GetHashCode()
        End Function


    End Class

End Namespace