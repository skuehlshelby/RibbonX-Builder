Imports RibbonX.Utilities

Namespace InternalApi
    Friend MustInherit Class RibbonProperty
        Private Sub New()
        End Sub

        Private NotInheritable Class DefaultProperty(Of T)
            Implements IRibbonProperty(Of T)

            Private ReadOnly value As T

            Public Sub New(name As IPropertyName, category As IPropertyCategory, value As T)
                Me.Name = NotNull(name)
                Me.Category = NotNull(category)
                Me.value = value
                IsReadOnly = True
                IsDefault = True
            End Sub

            Public ReadOnly Property Name As IPropertyName Implements IRibbonProperty.Name

            Public ReadOnly Property Category As IPropertyCategory Implements IRibbonProperty.Category

            Public ReadOnly Property IsReadOnly As Boolean Implements IRibbonProperty.IsReadOnly

            Public ReadOnly Property IsDefault As Boolean Implements IRibbonProperty.IsDefault

            Public Function GetValue() As T Implements IRibbonProperty(Of T).GetValue
                Return value
            End Function

            Public Function SetValue(value As T) As Boolean Implements IRibbonProperty(Of T).SetValue
                Throw New InvalidOperationException()
            End Function

            Public Function GetBoxedValue() As Object Implements IRibbonProperty.GetBoxedValue
                Return value
            End Function

            Public Overrides Function ToString() As String
                Return $"Default Property {Name} = {GetValue()}"
            End Function

            Public Function ToXml() As String Implements IRibbonProperty.ToXml
                Return String.Empty
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

            Public Function Clone() As Object Implements ICloneable.Clone
                Return New DefaultProperty(Of T)(Name, Category, value)
            End Function

        End Class

        Private NotInheritable Class ReadonlyProperty(Of T)
            Implements IRibbonProperty(Of T)

            Private ReadOnly value As T

            Public Sub New(name As IPropertyName, category As IPropertyCategory, value As T)
                Me.Name = NotNull(name)
                Me.Category = NotNull(category)
                Me.value = value
                IsReadOnly = True
                IsDefault = False
            End Sub

            Public ReadOnly Property Name As IPropertyName Implements IRibbonProperty.Name

            Public ReadOnly Property Category As IPropertyCategory Implements IRibbonProperty.Category

            Public ReadOnly Property IsReadOnly As Boolean Implements IRibbonProperty.IsReadOnly

            Public ReadOnly Property IsDefault As Boolean Implements IRibbonProperty.IsDefault

            Public Function GetValue() As T Implements IRibbonProperty(Of T).GetValue
                Return value
            End Function

            Public Function SetValue(value As T) As Boolean Implements IRibbonProperty(Of T).SetValue
                Throw New InvalidOperationException()
            End Function

            Public Function GetBoxedValue() As Object Implements IRibbonProperty.GetBoxedValue
                Return value
            End Function

            Public Overrides Function ToString() As String
                Return $"Readonly Property {Name} = {GetValue()}"
            End Function

            Public Function ToXml() As String Implements IRibbonProperty.ToXml
                Return $"{Name.ToXml()}=""{If(TypeOf value Is Boolean, value.CamelCase(), value.ToString())}"""
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

            Public Function Clone() As Object Implements ICloneable.Clone
                Return New DefaultProperty(Of T)(Name, Category, value)
            End Function

        End Class

        Friend Class MutableProperty(Of T)
            Implements IRibbonProperty(Of T)

            Private value As T
            Private ReadOnly callback As String

            Public Sub New(name As IPropertyName, category As IPropertyCategory, value As T, callback As String)
                Me.Name = NotNull(name)
                Me.Category = NotNull(category)
                Me.callback = NotNullOrWhiteSpace(callback)
                Me.value = value
                IsReadOnly = False
                IsDefault = False
            End Sub

            Public ReadOnly Property Name As IPropertyName Implements IRibbonProperty.Name

            Public ReadOnly Property Category As IPropertyCategory Implements IRibbonProperty.Category

            Public ReadOnly Property IsReadOnly As Boolean Implements IRibbonProperty.IsReadOnly

            Public ReadOnly Property IsDefault As Boolean Implements IRibbonProperty.IsDefault

            Public Function GetBoxedValue() As Object Implements IRibbonProperty.GetBoxedValue
                Return value
            End Function

            Public Function GetValue() As T Implements IRibbonProperty(Of T).GetValue
                Return value
            End Function

            Public Function SetValue(value As T) As Boolean Implements IRibbonProperty(Of T).SetValue
                If Not Equals(Me.value, value) Then
                    Me.value = value
                    Return True
                End If

                Return False
            End Function

            Public Overrides Function Equals(obj As Object) As Boolean
                Return Equals(TryCast(obj, IRibbonProperty))
            End Function

            Public Overloads Function Equals(other As IRibbonProperty) As Boolean Implements IEquatable(Of IRibbonProperty).Equals
                Return other IsNot Nothing AndAlso Category.Equals(other.Category)
            End Function

            Public Overrides Function ToString() As String
                Return $"Mutable Property {Name} = {GetValue()}"
            End Function

            Public Function ToXml() As String Implements IRibbonProperty.ToXml
                Return $"{Name.ToXml()}=""{callback}"""
            End Function

            Public Overrides Function GetHashCode() As Integer
                Return Category.GetHashCode()
            End Function

            Public Function Clone() As Object Implements ICloneable.Clone
                Return New MutableProperty(Of T)(Name, Category, value, callback)
            End Function
        End Class

        Public Shared Function Create(Of T)(name As IPropertyName, category As IPropertyCategory, value As T, Optional isDefault As Boolean = False) As IRibbonProperty(Of T)
            If isDefault Then
                Return New DefaultProperty(Of T)(name, category, value)
            Else
                Return New ReadonlyProperty(Of T)(name, category, value)
            End If
        End Function

        Public Shared Function Create(Of T)(name As IPropertyName, category As IPropertyCategory, value As T, callback As String) As IRibbonProperty(Of T)
            Return New MutableProperty(Of T)(name, category, value, callback)
        End Function

    End Class

End Namespace