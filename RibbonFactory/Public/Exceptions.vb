
Imports System.Runtime.Serialization
Imports RibbonX.Api

Public NotInheritable Class MissingOnLoadException
    Inherits Exception
    Implements ISerializable

    Public Sub New()
        MyBase.New(String.Join(" ", "Cannot update the Office host UI without an IRibbonUI object.",
                               "This is provided by the Office host via the OnLoad callback.",
                               $"To register an OnLoad callback, use {NameOf(IRibbonBuilder)}.{NameOf(IRibbonBuilder.OnLoad)}"))
    End Sub

    Public Sub New(info As SerializationInfo, context As StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class

Public NotInheritable Class PropertyNotSettableException
    Inherits Exception
    Implements ISerializable

    Private Const MESSAGE_TEMPLATE As String = "Property '{0}' is not settable. Are you missing a callback for this property? " &
                                               "A callback method must be provided which tells your Connection class (the one " &
                                               "which implements IDTExtensibility2 or IRibbonExtensibility) how to set this " &
                                               "property. This is typically done in a method like with{0}(value, callback) on the " &
                                               "I_ControlType_Builder which corresponds to this class."

    Public Sub New()
    End Sub

    Public Sub New(propertyName As String)
        MyBase.New(String.Format(MESSAGE_TEMPLATE, propertyName))
    End Sub

    Public Sub New(propertyName As String, innerException As Exception)
        MyBase.New(String.Format(MESSAGE_TEMPLATE, propertyName), innerException)
    End Sub

    Public Sub New(info As SerializationInfo, context As StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class

Public NotInheritable Class NoSuchPropertyException

End Class

Public NotInheritable Class UnreachableEventException
    Inherits Exception
    Implements ISerializable

    Private Const MESSAGE_TEMPLATE As String = "Attempted to register a handler for event '{0}', but it cannot be triggered. " &
                                               "A callback method must be provided which tells your Connection class (the one " &
                                               "which implements IDTExtensibility2 or IRibbonExtensibility) how to trigger this " &
                                               "event when change events are received from the Office host application. This is " &
                                               "done in the On_EventName_(callback...) method on the I_ControlType_Builder which " &
                                               "corresponds to this class."

    Public Sub New()
    End Sub

    Public Sub New(eventName As String)
        MyBase.New(String.Format(MESSAGE_TEMPLATE, eventName))
    End Sub

    Public Sub New(eventName As String, innerException As Exception)
        MyBase.New(String.Format(MESSAGE_TEMPLATE, eventName), innerException)
    End Sub

    Public Sub New(info As SerializationInfo, context As StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class