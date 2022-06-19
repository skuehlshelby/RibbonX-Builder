Imports RibbonX.Utilities

Namespace Controls.Attributes.Internal

    Friend Class RibbonAttributeReadWrite(Of T)
        Implements IRibbonAttributeReadWrite(Of T)
        Implements ICloneable

        Public Event ValueChanged() Implements IRibbonAttribute.ValueChanged

        Private Const XmlTemplate As String = "{0}=""{1}"""

        Private ReadOnly _callback As String
        Private _value As T


        Public Sub New(value As T, name As AttributeName, category As AttributeCategory, callback As String)
            Require(Of ArgumentNullException)(name IsNot Nothing)
            Require(Of ArgumentNullException)(category IsNot Nothing)
            Require(Of ArgumentException)(Not String.IsNullOrEmpty(callback))

            _value = value
            _callback = callback
            Me.Name = name
            Me.Category = category
        End Sub

        Public ReadOnly Property Name As AttributeName Implements IRibbonAttribute.Name

        Public ReadOnly Property Category As AttributeCategory Implements IRibbonAttribute.Category

        Public Function GetValue() As T Implements IRibbonAttributeReadOnly(Of T).GetValue
            Return _value
        End Function

        Public Sub SetValue(value As T) Implements IRibbonAttributeReadWrite(Of T).SetValue
            If Not _value.Equals(value) Then
                _value = value
                RaiseEvent ValueChanged()
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(XmlTemplate, _Name.CamelCase(), _callback)
        End Function

        Private Function Clone() As Object Implements ICloneable.Clone
            Dim valueCopy As T = If(TypeOf _value Is ICloneable, DirectCast(DirectCast(_value, ICloneable).Clone(), T), _value)

            Return New RibbonAttributeReadWrite(Of T)(valueCopy, Name.Clone(), Category.Clone(), String.Copy(_callback))
        End Function

    End Class

End Namespace