Imports System.Linq.Expressions
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports RibbonX.Controls
Imports RibbonX.Controls.Base
Imports RibbonX.Testing

Friend Module Extensions

    <Extension()>
    Public Sub ValueChangedIsRaisedOnce(Of T As RibbonElement)(assert As Assert, element As T, action As Action(Of T))
        Dim timesRaised As Integer = 0

        AddHandler element.ValueChanged, Sub(o, e) timesRaised += 1

        action.Invoke(element)

        Assert.AreEqual(1, timesRaised)
    End Sub

    <Extension()>
    Public Sub NoPropertiesAreNull(Of T As RibbonElement)(assert As Assert, control As T)
        For Each propertyInfo As PropertyInfo In GetType(T).GetProperties()
            If Not String.Equals(propertyInfo.Name, NameOf(RibbonElement.Tag)) Then
                Assert.IsNotNull(propertyInfo.GetValue(control), $"Property '{propertyInfo.Name}' on {control.ID} is null.")
            End If
        Next
    End Sub

    <Extension()>
    Public Sub PropertiesCannotBeModified(Of T As RibbonElement)(assert As Assert, control As T)
        Dim isReadWrite As Func(Of PropertyInfo, Boolean) = Function(p) p.CanRead AndAlso p.CanWrite
        Dim properties As IEnumerable(Of PropertyInfo) = GetType(T).GetProperties().Where(isReadWrite)

        For Each [property] As PropertyInfo In properties
            If Not String.Equals([property].Name, NameOf(RibbonElement.Tag)) Then
                Assert.ThrowsException(Of TargetInvocationException)(Sub() [property].SetValue(control, Nothing))
            End If
        Next
    End Sub

    <Extension()>
    Public Sub SharedPropertiesAreEqual(Of T As RibbonElement, K As RibbonElement)(assert As Assert, control As T, otherControl As K)
        Dim isReadWrite As Func(Of PropertyInfo, Boolean) = Function(p) p.CanRead AndAlso p.CanWrite
        Dim controlProperties As IEnumerable(Of PropertyInfo) = control.GetType().GetProperties().Where(isReadWrite)
        Dim otherControlProperties As IEnumerable(Of PropertyInfo) = otherControl.GetType().GetProperties().Where(isReadWrite)
        Dim controlPropertyNames As IEnumerable(Of String) = controlProperties.Select(Function(p) p.Name)
        Dim otherControlPropertyNames As IEnumerable(Of String) = otherControlProperties.Select(Function(p) p.Name)
        Dim namesOfSharedProperties As IEnumerable(Of String) = controlPropertyNames.Intersect(otherControlPropertyNames)

        For Each name As String In namesOfSharedProperties
            Dim firstControlProperty As PropertyInfo = controlProperties.Single(Function(p) p.Name.Equals(name))
            Dim secondControlProperty As PropertyInfo = otherControlProperties.Single(Function(p) p.Name.Equals(name))
            Dim errorMessage As String = $"Value of property '{name}' on {control.ID} and {otherControl.ID} do not match. {control.ID}: {firstControlProperty.GetValue(control)} vs. {secondControlProperty.GetValue(otherControl)}"
            Assert.AreEqual(firstControlProperty.GetValue(control), secondControlProperty.GetValue(otherControl), errorMessage)
        Next
    End Sub

    <Extension()>
    Public Sub ValidRibbonXIsProduced(Of T As RibbonElement)(assert As Assert, control As T)
        Dim ribbon As Ribbon = Nothing

        If TypeOf control Is Group Then
            ribbon = New Ribbon(New Tab(DirectCast(CType(control, Object), Group)))
        ElseIf TypeOf control Is Tab Then
            ribbon = New Ribbon(DirectCast(CType(control, Object), Tab))
        Else
            ribbon = New Ribbon(New Tab(New Group(controls:={control})))
        End If

        Dim errors As RibbonErrorLog = ribbon.GetErrors()

        Assert.IsTrue(ribbon.GetErrors().None, String.Join(Environment.NewLine, errors.Errors))

        Debug.WriteLine(ribbon.RibbonX)
    End Sub

    <Extension()>
    Public Sub PropertyMayBeModified(Of TElement As RibbonElement, TProperty)(assert As Assert, control As TElement, propertyExpression As Expression(Of Func(Of TElement, TProperty)), updatedValue As TProperty)
        Dim memberExpression As MemberExpression = DirectCast(propertyExpression.Body, MemberExpression)
        Dim propertyInfo As PropertyInfo = CType(memberExpression.Member, PropertyInfo)
        Dim currentValue As TProperty = CType(propertyInfo.GetValue(control), TProperty)
        Assert.AreNotEqual(currentValue, updatedValue)
        propertyInfo.SetValue(control, updatedValue)
        currentValue = CType(propertyInfo.GetValue(control), TProperty)
        Assert.AreEqual(currentValue, updatedValue)
    End Sub
End Module
