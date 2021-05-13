Imports System.IO
Imports System.Text
Imports NetOffice.OfficeApi
Imports RibbonFactory.Builders
Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls

<TestClass()>
Public Class SimpleRibbonTest
    Inherits RibbonTestBase
    Private _ribbon As Ribbon

    <TestMethod()>
    Public Sub BuildSimple()
        
        Try
            RedirectConsoleOutput()
            
            Dim button As Button = New ButtonBuilder() _
                    .WithLabel("My Button") _
                    .WithScreenTip("My favorite button.") _
                    .ThatDoes(AddressOf OnAction, Sub() MsgBox("Hello World!")) _
                    .Build()
        
            Dim group As Group = New GroupBuilder().WithLabel("Test Group").Build()
        
            Dim tab As Tab = New TabBuilder().WithLabel("Test Tab").Build()
        
            group.Add(button)
        
            tab.AddGroup(group)
        
            _ribbon = New Ribbon(startFromScratch:= False, [nameSpace]:= "Testing")
        
            _ribbon.Tabs.Add(tab)
                                                    
            dim unused As String = _ribbon.Build()
        Finally
            RestoreOriginalConsoleOutput()
        End Try
        
        Assert.AreEqual(string.Empty, GetContentsOfRedirectedConsoleOutput())

    End Sub
    
End Class