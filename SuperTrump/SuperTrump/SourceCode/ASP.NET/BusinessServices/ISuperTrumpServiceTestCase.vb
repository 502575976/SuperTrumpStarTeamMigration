Imports NUnit.Framework

<TestFixture()> _
Public Class ISuperTrumpServiceTestCase

    <Test()> _
            Public Sub TestConvertPRMToXML()
        Dim obj As New ISuperTrumpService
        Dim objInput, objOutPut As New Xml.XmlDocument
        Try
            objInput.Load("XML_IN\ConvertPRMToXMLtest.xml")
            objOutPut.Load("XML_OUT\ConvertPRMToXMLtest.xml")
            Assert.Greater(objOutPut.InnerText, obj.ConvertPRMToXML(objInput.InnerXml))
        Catch ex As Exception
            Throw ex
        Finally
            obj.Dispose()
            objInput = Nothing
            objOutPut = Nothing
        End Try
    End Sub

    <Test()> _
           Public Sub TestGeneratePRMFiles()
        Dim obj As New ISuperTrumpService
        Dim objInput, objOutPut As New Xml.XmlDocument
        Try
            objInput.Load("XML_IN\GeneratePRMFilestest.xml")
            objOutPut.Load("XML_OUT\GeneratePRMFilestest.xml")
            Assert.Greater(objOutPut.InnerText, obj.GeneratePRMFiles(objInput.InnerXml))
        Catch ex As Exception
            Throw ex
        Finally
            obj.Dispose()
            objInput = Nothing
            objOutPut = Nothing
        End Try
    End Sub

    <Test()> _
          Public Sub TestGetAmortizationSchedule()
        Dim obj As New ISuperTrumpService
        Dim objInput, objOutPut As New Xml.XmlDocument
        Try
            objInput.Load("XML_IN\GetAmortizationScheduletest.xml")
            objOutPut.Load("XML_OUT\GetAmortizationScheduletest.xml")
            Assert.Greater(objOutPut.InnerText, obj.GetAmortizationSchedule(objInput.InnerXml))
        Catch ex As Exception
            Throw ex
        Finally
            obj.Dispose()
            objInput = Nothing
            objOutPut = Nothing
        End Try
    End Sub

    <Test()> _
          Public Sub TestGetPricingReports()
        Dim obj As New ISuperTrumpService
        Dim objInput, objOutPut As New Xml.XmlDocument
        Try
            objInput.Load("XML_IN\GetPricingReportstest.xml")
            objOutPut.Load("XML_OUT\GetPricingReportstest.xml")
            Assert.Greater(objOutPut.InnerText, obj.GetPricingReports(objInput.InnerXml))
        Catch ex As Exception
            Throw ex
        Finally
            obj.Dispose()
            objInput = Nothing
            objOutPut = Nothing
        End Try
    End Sub

    <Test()> _
          Public Sub TestGetPRMParams()
        Dim obj As New ISuperTrumpService
        Dim objInput, objOutPut As New Xml.XmlDocument
        Try
            objInput.Load("XML_IN\GetPRMParamstest.xml")
            objOutPut.Load("XML_OUT\GetPRMParamstest.xml")
            Assert.Greater(objOutPut.InnerText, obj.GetPRMParams(objInput.InnerXml))
        Catch ex As Exception
            Throw ex
        Finally
            obj.Dispose()
            objInput = Nothing
            objOutPut = Nothing
        End Try
    End Sub

    <Test()> _
         Public Sub TestModifyPRMFiles()
        Dim obj As New ISuperTrumpService
        Dim objInput, objOutPut As New Xml.XmlDocument
        Try
            objInput.Load("XML_IN\ModifyPRMFilestest.xml")
            objOutPut.Load("XML_OUT\ModifyPRMFilestest.xml")
            Assert.Greater(objOutPut.InnerText, obj.ModifyPRMFiles(objInput.InnerXml))
        Catch ex As Exception
            Throw ex
        Finally
            obj.Dispose()
            objInput = Nothing
            objOutPut = Nothing
        End Try
    End Sub

End Class
