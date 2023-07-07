Imports System.Xml.Schema
Imports RibbonX.Utilities

Namespace Testing

    ''' <summary>
    ''' <para>
    ''' Use this class to validate a complete ribbon after it has been created.<br/>
    ''' This class validates the supplied ribbonX string in its constructor using<br/>
    ''' the supplied <c>XmlSchemas</c>. If you do not have any CustomUI schemas,<br/>
    ''' the 2009 schema is available as part of this library and can be found by<br/>
    ''' calling the static function Ribbon.GetCustomUIVersion2009().
    ''' </para>
    ''' 
    ''' <para>
    ''' Use this class in your unit tests to verify that there are no errors in your ribbon.<br/>
    ''' In the event that errors are found, the Errors property will display all errors. The<br/>
    ''' standard behavior of MS Office applications is to display only the first error.<br/>
    ''' Using MS Office to check for ribbon errors is not advised.
    ''' </para>
    ''' </summary>
    Public NotInheritable Class RibbonErrorLog
        Private ReadOnly _errors As ICollection(Of String)

        ''' <summary>
        ''' Validates the supplied <paramref name="ribbonX"/>, using the supplied <paramref name="xmlSchemas"/>.<br/>
        ''' The properties of this class facilitate querying the validation results.
        ''' </summary>
        ''' <param name="ribbonX">The ribbonX. Obtained from a <c>Ribbon</c> object.</param>
        ''' <param name="xmlSchemas">The <c>XmlSchema</c>s that should be used for validation.</param>
        Public Sub New(ribbonX As String, ParamArray xmlSchemas As XmlSchema())
            Require(Of ArgumentNullException)(xmlSchemas IsNot Nothing, "An XML schema is required in order to validate the ribbon.")
            Require(Of ArgumentException)(xmlSchemas.Any(), "At least one XML schema must be supplied.")

            _errors = New List(Of String)
            ValidateRibbon(ribbonX, xmlSchemas)
        End Sub

        ''' <summary>
        ''' Are there no errors?
        ''' </summary>
        ''' <returns><c>True</c> if the supplied XML was error-free.</returns>
        Public ReadOnly Property None As Boolean
            Get
                Return Not _errors.Any()
            End Get
        End Property

        ''' <summary>
        ''' The errors in the supplied XML.
        ''' </summary>
        ''' <returns>The errors, or an empty collection.</returns>
        Public ReadOnly Property Errors As IEnumerable(Of String)
            Get
                Return _errors
            End Get
        End Property

        Private Sub ValidateRibbon(ribbonX As String, ParamArray xmlSchemas As XmlSchema())
            Dim xsd As XmlSchemaSet = New XmlSchemaSet()

            Array.ForEach(xmlSchemas, Sub(schema) xsd.Add(schema))

            XDocument.Parse(ribbonX).Validate(xsd, AddressOf LogXmlError)
        End Sub

        ''' <summary>
        ''' Records parsing errors in the format '{NodeName} {Severity}: {Error Message}'
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub LogXmlError(sender As Object, e As ValidationEventArgs)
            If TypeOf sender Is XAttribute Then
                _errors.Add($"{e.Severity.ToString().ToUpper}: {e.Message} for the element '{DirectCast(sender, XAttribute).Parent.Name.LocalName}'")
            Else
                _errors.Add($"{e.Severity.ToString().ToUpper}: {e.Message}'")
            End If
        End Sub
    End Class
End Namespace