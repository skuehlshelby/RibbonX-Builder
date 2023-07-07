Namespace Images.BuiltIn

    Public Class XML
        Inherits ImageMSO

        Private Sub New(value As Integer, name As String)
            MyBase.New(value, name)
        End Sub

        Public Shared ReadOnly Property DesignXml As XML = New XML(1, NameOf(DesignXml))
        Public Shared ReadOnly Property ExportXmlFile As XML = New XML(2, NameOf(ExportXmlFile))
        Public Shared ReadOnly Property ImportXmlFile As XML = New XML(3, NameOf(ImportXmlFile))
        Public Shared ReadOnly Property GroupXml As XML = New XML(4, NameOf(GroupXml))
        Public Shared ReadOnly Property XmlDataRefresh As XML = New XML(5, NameOf(XmlDataRefresh))
        Public Shared ReadOnly Property XmlEditQuery As XML = New XML(6, NameOf(XmlEditQuery))
        Public Shared ReadOnly Property XmlExpansionPacksExcel As XML = New XML(7, NameOf(XmlExpansionPacksExcel))
        Public Shared ReadOnly Property XmlExpansionPacksWord As XML = New XML(8, NameOf(XmlExpansionPacksWord))
        Public Shared ReadOnly Property XmlExport As XML = New XML(9, NameOf(XmlExport))
        Public Shared ReadOnly Property XmlImport As XML = New XML(10, NameOf(XmlImport))
        Public Shared ReadOnly Property XmlMapProperties As XML = New XML(11, NameOf(XmlMapProperties))
        Public Shared ReadOnly Property XmlSchema As XML = New XML(12, NameOf(XmlSchema))
        Public Shared ReadOnly Property XmlSource As XML = New XML(13, NameOf(XmlSource))
        Public Shared ReadOnly Property XmlStructure As XML = New XML(14, NameOf(XmlStructure))
        Public Shared ReadOnly Property XmlTransformation As XML = New XML(15, NameOf(XmlTransformation))


        Public Overrides Function Clone() As Object
            Return New XML(value, name)
        End Function

    End Class

End Namespace