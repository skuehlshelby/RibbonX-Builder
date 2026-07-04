Imports System.ComponentModel
Imports System.Runtime.CompilerServices

Namespace InternalApi
    Friend Interface IPropertyName
        Inherits IEquatable(Of IPropertyName)
        Inherits ICloneable
        Function ToXml() As String
    End Interface

    Friend Interface IPropertyCategory
        Inherits IReadOnlyCollection(Of IPropertyName)
        Inherits IEquatable(Of IPropertyCategory)
        Inherits ICloneable
    End Interface

    Friend Interface IPropertyCategory(Of Out T)
        Inherits IPropertyCategory
    End Interface

    Friend Interface IRibbonProperty
        Inherits IEquatable(Of IRibbonProperty)
        Inherits ICloneable
        ReadOnly Property Name As IPropertyName
        ReadOnly Property Category As IPropertyCategory
        ReadOnly Property IsReadOnly As Boolean
        ReadOnly Property IsDefault As Boolean
        Function GetBoxedValue() As Object
        Function ToXml() As String
    End Interface

    Friend Interface IRibbonProperty(Of T)
        Inherits IRibbonProperty

        Function GetValue() As T
        Function SetValue(value As T) As Boolean
    End Interface

    Friend Interface IPropertyCollection
        Inherits INotifyPropertyChanged
        Inherits ICloneable
        Inherits ICollection(Of IRibbonProperty)
        Function [Get](Of T)(category As IPropertyCategory(Of T)) As T
        Sub [Set](Of T)(value As T, category As IPropertyCategory(Of T), <CallerMemberName()> Optional caller As String = vbNullString)
        Function Raw(category As IPropertyCategory) As IRibbonProperty
        Function Raw(ParamArray categories As IPropertyCategory()) As IRibbonProperty()
        Sub Forward(from As IPropertyCategory, [to] As IPropertyCategory)
        Sub [AddHandler](eventName As String, handler As [Delegate])
        Sub [RemoveHandler](eventName As String, handler As [Delegate])
        Sub [RaiseEvent](Of TEventArgs As EventArgs)(eventName As String, sender As Object, e As TEventArgs)
        Sub Merge(other As IPropertyCollection)
        Function ToXml(Optional exclude As ExcludedAttributes = ExcludedAttributes.None) As String
    End Interface
End Namespace