Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Configuration
Imports System.Reflection
Imports System.Web.Caching
Imports SupertrumpService.BusinessServices
<WebServiceBinding(Name:="ISuperTrumpServiceSoapPort"), _
 WebServiceBinding(Name:="ISuperTrumpServiceSoapBinding", _
                   Namespace:="ISuperTrumpServiceSoapBinding"), _
 WebServiceBinding(Name:="IClientServiceSoapPort"), _
WebServiceBinding(Name:="IClientServiceSoapBinding", _
                   Namespace:="IClientServiceSoapBinding")> _
Public Class SupertrumpService
    Inherits System.Web.Services.WebService
#Region "IClientService"
    <SoapDocumentMethod(Binding:="IClientServiceSoapPort"), STAThreadAttribute(), WebMethod()> _
       Public Function Ping() As String

        Dim lobjSTWebSuperTrumpService = ISuperTrumpService.Instance
        Dim lobjSTWebClientService = IClientService.Instance
        Try
            Return lobjSTWebSuperTrumpService.Ping() & "Thread State:- ISuperTrumpService : " & lobjSTWebSuperTrumpService.GetThradApartment & " IClientService : " & lobjSTWebClientService.GetThradApartment
        Catch ex As Exception
            Return ex.Message
        Finally
            lobjSTWebClientService = Nothing
            lobjSTWebSuperTrumpService = Nothing
        End Try
    End Function
    '================================================================
    'MODULE  : IClientService
    'PURPOSE : This interface provides all customized methods for the
    '          Client applications. These methods internally call the
    '          methods in the ISuperTrumpService interface.
    '================================================================

    '================================================================
    'METHOD  : ProcessMQMessage
    'PURPOSE : To process messages sent asynchronously by Client
    '          Applications through MQ Series.
    'PARMS   :
    '          astrMQMsgInfoXML [String] = XML string containing
    '          the Message Id, data, reply queue manager name, reply
    '          queue name and Correlation Id.
    '
    '          Sample Input parameter structure:
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <MQ_MESSAGE_INFO>
    '                <MESSAGE_ID>214D512053494542454C50524F4420204464C33C73201F00</MESSAGE_ID>
    '                <MESSAGE_DATA><![CDATA[<MY_MSG>...</MY_MSG>]]></MESSAGE_DATA>
    '                <REPLY_QUEUE_MANAGER>MY_Q_MGR</REPLY_QUEUE_MANAGER>
    '                <REPLY_QUEUE>MY_Q</REPLY_QUEUE>
    '                <CORRELATION_ID>414D512053494542454C50524F4420204464C33C73201F00</CORRELATION_ID>
    '            </MQ_MESSAGE_INFO>
    'RETURN  : String = XML string containing Instructions to MQ.
    '================================================================
    <SoapDocumentMethod(Binding:="IClientServiceSoapPort"), STAThreadAttribute(), WebMethod()> _
    Public Function ProcessMQMessage(ByVal astrMQMsgInfoXML As String) As String

        Dim obj = IClientService.Instance
        Try
            Return obj.ProcessMQMessage(astrMQMsgInfoXML)
        Catch ex As Exception
            Return ex.Message()
        Finally
            obj = Nothing
        End Try

    End Function
    '================================================================
    'METHOD  : ProcessPricingRequest
    'PURPOSE : To process the Pricing Request XML
    'PARMS   :
    '          astrPricingRequestXML [String] = Pricing Request XML
    'RETURN  : String = Pricing Response XML
    '================================================================
    <SoapDocumentMethod(Binding:="IClientServiceSoapPort"), WebMethod()> _
    Public Function ProcessPricingRequest(ByVal astrPricingRequestXML As String) As String

        Dim lobjSTWebService = IClientService.Instance
        Try
            If (HttpRuntime.Cache.Get("STObjectByCache") Is Nothing) Then
                lobjSTWebService = IClientService.Instance
                HttpRuntime.Cache.Insert("STObjectByCache", lobjSTWebService)
            Else
                lobjSTWebService = HttpRuntime.Cache.Get("STObjectByCache")
            End If
            Return lobjSTWebService.ProcessPricingRequest(astrPricingRequestXML)
        Catch ex As Exception
            Return ex.Message()
        Finally
            lobjSTWebService = Nothing
        End Try
    End Function
#End Region
#Region "ISuperTrumpService Code Area"

    Private Const cADHOC_QUERY_RESULT_XML As String = "<PRM_INFO><PRM_FILE><AD_HOC_QUERY></AD_HOC_QUERY></PRM_FILE></PRM_INFO>"

    '================================================================
    'METHOD  : ConvertPRMToXML
    'PURPOSE : To convert the binary PRM file(s) to their XML
    '          equivalent.
    'PARMS   :
    '          astrPRMFileListXML [String] = XML string containing
    '          the binary PRM files. This XML will conform to the
    '          PRMFileListXML.xsd schema.
    '
    '          Sample Input Parameter structure:
    '           <PRM_FILE_LIST>
    '               <PRM_FILE>
    '                   <FILE_NAME>…</FILE_NAME>
    '                   <FILE_DATA>…</FILE_ DATA>
    '               </PRM_FILE>
    '               <PRM_FILE>
    '                   <FILE_NAME>…</FILE_NAME>
    '                   <FILE_DATA>…</FILE_ DATA>
    '               </PRM_FILE>
    '               …
    '           </PRM_FILE_LIST>
    'RETURN  : String = An XML string containing XML equivalent for
    '          each PRM File. It will also contain an error message
    '          for each erroneous PRM File.
    '
    '          Sample Return XML structure:
    '           <PRM_FILE_LIST>
    '               <PRM_FILE>
    '                   <FILE_NAME>…</FILE_NAME>
    '                   <PRM_XML>…</PRM_XML>
    '               </PRM_FILE>
    '               <PRM_FILE>
    '                   <FILE_NAME>…</FILE_NAME>
    '                   <ERROR>
    '                       <ERROR_NBR>…</ERROR_NBR>
    '                       <ERROR_DESC>…</ERROR_DESC>
    '                   </ERROR>
    '               </PRM_FILE>
    '               …
    '           </PRM_FILE_LIST>
    '
    '           OR in case of application error
    '
    '           <PRM_FILE_LIST>
    '               <ERROR>
    '                   <ERROR_NBR>…</ERROR_NBR>
    '                   <ERROR_DESC>…</ERROR_DESC>
    '               </ERROR>
    '           </PRM_FILE_LIST>
    '================================================================
    <SoapDocumentMethod(Binding:="ISuperTrumpServiceSoapPort"), STAThreadAttribute(), WebMethod()> _
    Function ConvertPRMToXML(ByVal astrPRMFileListXML As String) As String

        Dim lobjSTWebService = ISuperTrumpService.Instance
        Try
            Return lobjSTWebService.ConvertPRMToXML(astrPRMFileListXML)
        Catch ex As Exception
            Return ex.Message()
        Finally
            lobjSTWebService = Nothing
        End Try
    End Function
    '================================================================
    'METHOD  : GeneratePRMFiles
    'PURPOSE : To generate binary PRM file for each set of PRM
    '          parameters and Meta data.
    'PARMS   :
    '          astrPRMInfoXML [String] = XML string containing the
    '          PRM Parameters and Meta data required to generate the
    '          binary PRM file(s).
    '
    '          Sample Input Parameter structure:
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <PRM_INFO>
    '                <PRM_FILE>
    '                    <PRM_META_DATA>
    '                        <FILE_NAME>MyPRMFile.prm</FILE_NAME>
    '                        <TEMPLATE_NAME>USA 5 MACRS.TEM</TEMPLATE_NAME>
    '                        <MODE>Lessor</MODE>
    '                    </PRM_META_DATA>
    '                    <PRM_PARAMS>
    '                        <TRANSACTIONAMOUNT>25000000</TRANSACTIONAMOUNT>
    '                        <TRANSACTIONSTARTDATE>2002-08-20</TRANSACTIONSTARTDATE>
    '                        <RESIDUALAMOUNT>100000</RESIDUALAMOUNT>
    '                        <NUMBEROFPAYMENTS>60</NUMBEROFPAYMENTS>
    '                        <PERIODICITY>Monthly</PERIODICITY>
    '                        <PAYMENTTIMING>Advance</PAYMENTTIMING>
    '                        <STRUCTURE>Level</STRUCTURE>
    '                        <TARGETDATA>
    '                            <TYPEOFSTATISTIC>Yield</TYPEOFSTATISTIC>
    '                            <STATISTICINDEX>1</STATISTICINDEX>
    '                            <NEPA>Pre-tax nominal</NEPA>
    '                            <TARGETVALUE>0.075</TARGETVALUE>
    '                            <ADJUST>Rent</ADJUST>
    '                            <ADJUSTMENTMETHOD>Proportional</ADJUSTMENTMETHOD>
    '                        </TARGETDATA>
    '                    </PRM_PARAMS>
    '                </PRM_FILE>
    '                <PRM_FILE>
    '                    <PRM_META_DATA>
    '                        <FILE_NAME>ErrorPRMFile.prm</FILE_NAME>
    '                        ...
    '                    </PRM_META_DATA>
    '                    …
    '                </PRM_FILE>
    '                …
    '            </PRM_INFO>
    'RETURN  : String= XML string containing, the binary PRM File or
    '          <ERROR> tag, for each set of PRM Input Parameters.
    '          It may also return an <ERROR> tag for any general
    '          failure condition.
    '
    '            Sample Return XML structure:
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <PRM_FILE_LIST>
    '
    '                <!-- Sucessful generation of PRM file -->
    '                <PRM_FILE>
    '                    <FILE_NAME>MyPRMFile.prm</FILE_NAME>
    '                    <FILE_DATA>/CQAGAAAAAAAAAAAAAAACAAAA3AAAAAAA…</FILE_DATA>
    '                </PRM_FILE>
    '
    '                <!-- Error generating PRM file -->
    '                <PRM_FILE>
    '                    <FILE_NAME>ErrorPRMFile.prm </FILE_NAME>
    '                    <ERROR>
    '                        <ERROR_NBR>-1072896682</ERROR_NBR>
    '                        <ERROR_DESC>Error!!!...</ERROR_DESC>
    '                    </ERROR>
    '                </PRM_FILE>
    '                …
    '            </PRM_FILE_LIST>
    '
    '            OR In case of general failure:
    '
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <PRM_FILE_LIST>
    '                <ERROR>
    '                    <ERROR_NBR>-1072896682</ERROR_NBR>
    '                    <ERROR_DESC>Error!!!...</ERROR_DESC>
    '                </ERROR>
    '            </PRM_FILE_LIST>
    '================================================================
    <SoapDocumentMethod(Binding:="ISuperTrumpServiceSoapPort"), STAThreadAttribute(), WebMethod()> _
    Public Function GeneratePRMFiles(ByVal astrPRMInfoXML As String) As String

        Dim lobjSTWebService = ISuperTrumpService.Instance
        Try
            Return lobjSTWebService.GeneratePRMFiles(astrPRMInfoXML)
        Catch ex As Exception
            Return ex.Message()
        Finally
            lobjSTWebService = Nothing
        End Try

    End Function
    '================================================================
    'METHOD  : GetAmortizationSchedule
    'PURPOSE : To get amortization schedule for the inputted binary
    '          PRM file(s).
    'PARMS   :
    '          astrPRMFileListXML [String] = XML string containing
    '          the List of binary PRM files.
    '
    '            Sample Input Parameter structure:
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <PRM_FILE_LIST>
    '                <PRM_FILE>
    '                    <FILE_NAME>LeasePRMFile.prm</FILE_NAME>
    '                    <FILE_DATA>/CQAGAAAAAAAAAAAAAAACAAAA3AAAAAAA…</FILE_DATA>
    '                </PRM_FILE>
    '                <PRM_FILE>
    '                    <FILE_NAME>ErrorPRMFile.prm</FILE_NAME>
    '                    <FILE_DATA>AAAAAAAAAAAAAAAAAAAAAPgADAP7…</FILE_DATA>
    '                </PRM_FILE>
    '                <PRM_FILE>
    '                    <FILE_NAME>LoanPRMFile.prm</FILE_NAME>
    '                    <FILE_DATA>M8R4KGxGuEAAAAAAAAAAAAAAAAAAAAAPgADAP7…</FILE_DATA>
    '                </PRM_FILE>
    '                …
    '            </PRM_FILE_LIST>
    '
    '            Note:
    '            1)  <FILE_NAME> tag must contain PRM file name with .prm extension.
    '            2)  <FILE_DATA> tag must contain binary value of type base64Binary.
    'RETURN  : String = XML string containing, the Rent Schedule data
    '          or a <ERROR> tag, for each binary PRM file. It may also
    '          return an <ERROR> tag for any general failure condition.
    '
    '            Sample Return XML structure:
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <AMORTIZATION_SCHEDULE_LIST>
    '
    '                <!-- Sucessful generation of amortization schedule -->
    '                <AMORTIZATION_SCHEDULE>
    '                    <PRM_FILE_NAME>LeasePRMFile.prm</PRM_FILE_NAME>
    '                    <PAYMENT_LIST>
    '                        <PAYMENT>
    '                            <PAYMENT_NUMBER>1</PAYMENT_NUMBER>
    '                            <PAYMENT_START_DATE>8/22/2002</PAYMENT_START_DATE>
    '                            <PAYMENT_AMOUNT>10000</PAYMENT_AMOUNT>
    '                            <LEASE_FACTOR>0.0543</LEASE_FACTOR>
    '                        </PAYMENT>
    '                        ...
    '                    </PAYMENT_LIST>
    '                </AMORTIZATION_SCHEDULE>
    '
    '                <!-- Error reading PRM file  -->
    '                <AMORTIZATION_SCHEDULE>
    '                    <PRM_FILE_NAME>ErrorPRMFile.prm </PRM_FILE_NAME>
    '                    <ERROR>
    '                        <ERROR_NBR>-1072896682</ERROR_NBR>
    '                        <ERROR_DESC>Error!!!...</ERROR_DESC>
    '                    </ERROR>
    '                </AMORTIZATION_SCHEDULE>
    '                ...
    '            </AMORTIZATION_SCHEDULE_LIST>
    '
    '            OR In case of general failure:
    '
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <AMORTIZATION_SCHEDULE_LIST>
    '                <ERROR>
    '                    <ERROR_NBR>-1072896682</ERROR_NBR>
    '                    <ERROR_DESC>Error!!!...</ERROR_DESC>
    '                </ERROR>
    '            </AMORTIZATION_SCHEDULE_LIST>
    '================================================================
    <SoapDocumentMethod(Binding:="ISuperTrumpServiceSoapPort"), STAThreadAttribute(), WebMethod()> _
    Public Function GetAmortizationSchedule(ByVal astrPRMFileListXML As String) As String

        Dim lobjSTWebService = ISuperTrumpService.Instance
        Try
            Return lobjSTWebService.GetAmortizationSchedule(astrPRMFileListXML)
        Catch ex As Exception
            Return ex.Message()
        Finally
            lobjSTWebService = Nothing
        End Try

    End Function
    '================================================================
    'METHOD  : GetPricingReports
    'PURPOSE : To get the specified report(s) for the inputted binary
    '          PRM file(s).
    'PARMS   :
    '          astrPricingRepInfoXML [String] = XML string containing
    '          the binary PRM files and information specifying what
    '          report(s) needs to be generated for each PRM file.
    '          This XML will conform to the PricingRepInfoXML.xsd
    '          schema.
    '
    '          Sample Input Parameter structure:
    '           <PRICING_REPORT_INFO>
    '               <PRICING_REPORT>
    '                   <PRM_FILE>
    '                       <FILE_NAME>…</FILE_NAME>
    '                       <FILE_DATA>…</FILE_ DATA>
    '                   </PRM_FILE>
    '                   <REPORT_LIST>
    '                       <REPORT_TYPE>…</REPORT_TYPE>
    '                       <REPORT_TYPE>…</REPORT_TYPE>
    '                       …
    '                   </REPORT_LIST>
    '               </PRICING_REPORT>
    '               …
    '           </PRICING_REPORT_INFO>
    'RETURN  : String = An XML string containing the pricing reports
    '          for each PRM File. It will also contain an error
    '          message for each erroneous PRM File and each pricing
    '          reports, which couldn't be generated.
    '
    '          Sample Return XML structure:
    '           <PRICING_REPORT_LIST>
    '               <PRICING_REPORT>
    '                   <PRM_FILE_NAME>…</ PRM_FILE_NAME>
    '                   <REPORT_LIST>
    '                       <REPORT>
    '                           <REPORT_TYPE>…</REPORT_TYPE>
    '                           <TEXT_REPORT>…</TEXT_REPORT>
    '                       </REPORT>
    '                       <REPORT>
    '                           <REPORT_TYPE>…</REPORT_TYPE>
    '                           <TEXT_REPORT>…</TEXT_REPORT>
    '                       </REPORT>
    '                       …
    '                   </REPORT_LIST>
    '               </PRICING_REPORT>
    '               <PRICING_REPORT>
    '                   <PRM_FILE_NAME>…</ PRM_FILE_NAME>
    '                   <ERROR>
    '                       <ERROR_NBR>…</ERROR_NBR>
    '                       <ERROR_DESC>…</ERROR_DESC>
    '                   </ERROR>
    '               </PRICING_REPORT>
    '               <PRICING_REPORT>
    '                   <PRM_FILE_NAME>…</ PRM_FILE_NAME>
    '                   <REPORT_LIST>
    '                       <REPORT>
    '                           <REPORT_TYPE>…</REPORT_TYPE>
    '                           <TEXT_REPORT>…</TEXT_REPORT>
    '                       </REPORT>
    '                       <REPORT>
    '                           <REPORT_TYPE>…</REPORT_TYPE>
    '                           <ERROR>
    '                               <ERROR_NBR>…</ERROR_NBR>
    '                               <ERROR_DESC>…</ERROR_DESC>
    '                           </ERROR>
    '                       </REPORT>
    '                       …
    '                   </REPORT_LIST>
    '               </PRICING_REPORT>
    '               …
    '           </PRICING_REPORT_LIST>
    '
    '           OR in case of application error
    '
    '           <PRICING_REPORT_LIST>
    '               <ERROR>
    '                   <ERROR_NBR>…</ERROR_NBR>
    '                   <ERROR_DESC>…</ERROR_DESC>
    '               </ERROR>
    '           </PRICING_REPORT_LIST>
    '================================================================
    <SoapDocumentMethod(Binding:="ISuperTrumpServiceSoapPort"), STAThreadAttribute(), WebMethod()> _
    Public Function GetPricingReports(ByVal astrPricingRepInfoXML As String) As String


        Dim lobjSTWebService = ISuperTrumpService.Instance
        Try

            Return lobjSTWebService.GetPricingReports(astrPricingRepInfoXML)
        Catch ex As Exception
            Return ex.Message()
        Finally
            lobjSTWebService = Nothing
        End Try

    End Function
    '================================================================
    'METHOD  : GetPRMParams
    'PURPOSE : To get specified PRM Parameters for the inputted
    '          binary PRM file(s).
    '          Note: This method is similar to the ConvertPRMToXML()
    '          method, but it will return only a subset of data than
    '          the one returned by the ConvertPRMToXML() method.
    'PARMS   :
    '          astrPRMParamsInfoXML [String]= XML string containing
    '          the List of PRM parameters.
    '
    '            Sample Input Parameter structure:
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <PRM_PARAMS_INFO>
    '                <PRM_PARAMS>
    '                    <PRM_PARAMS_SPECS>
    '                        <TRANSACTIONAMOUNT query="true"/>
    '                        <TRANSACTIONSTARTDATE query="true"/>
    '                        <RESIDUALAMOUNT query="true"/>
    '                        <STRUCTURE query="true"/>
    '                        <PERIODICITY query="true"/>
    '                        <PAYMENTTIMING query="true"/>
    '                        <NUMBEROFPAYMENTS query="true"/>
    '                        <TARGETDATA query="true"/>
    '                    </PRM_PARAMS_SPECS>
    '                    <PRM_FILE>
    '                        <FILE_NAME>LeasePRMFile.prm</FILE_NAME>
    '                        <FILE_DATA>/CQAGAAAAAAAAAAAAAAACAAAA3AAAAAAA…</FILE_DATA>
    '                    </PRM_FILE>
    '                </PRM_PARAMS>
    '                <PRM_PARAMS>
    '                    <PRM_PARAMS_SPECS>
    '                        <TRANSACTIONAMOUNT query="true"/>
    '                        <TRANSACTIONSTARTDATE query="true"/>
    '                        <STRUCTURE query="true"/>
    '                        <PERIODICITY query="true"/>
    '                        <PAYMENTTIMING query="true"/>
    '                        <NUMBEROFPAYMENTS query="true"/>
    '                        <TARGETDATA query="true"/>
    '                    </PRM_PARAMS_SPECS>
    '                    <PRM_FILE>
    '                        <FILE_NAME>LoanPRMFile.prm</FILE_NAME>
    '                        <FILE_DATA>M8R4KGxGuEAAAAAAAAAAAAAAAAAAAAAPgADAP7…</FILE_DATA>
    '                    </PRM_FILE>
    '                </PRM_PARAMS>
    '                <PRM_PARAMS>
    '                    <PRM_PARAMS_SPECS>
    '                        <TRANSACTIONAMOUNT query="true"/>
    '                        <TRANSACTIONSTARTDATE query="true"/>
    '                        <STRUCTURE query="true"/>
    '                        <PERIODICITY query="true"/>
    '                        <PAYMENTTIMING query="true"/>
    '                        <NUMBEROFPAYMENTS query="true"/>
    '                        <TARGETDATA query="true"/>
    '                    </PRM_PARAMS_SPECS>
    '                    <PRM_FILE>
    '                        <FILE_NAME>ErrorPRMFile.prm</FILE_NAME>
    '                        <FILE_DATA>AAAAAAAAAAAAAAAAAAAAAPgADAP7…</FILE_DATA>
    '                    </PRM_FILE>
    '                </PRM_PARAMS>
    '                …
    '            </PRM_PARAMS_INFO>
    '
    '            Note:
    '            1)  <FILE_NAME> tag must contain PRM file name with .prm extension.
    '            2)  <FILE_DATA> tag must contain binary value of type base64Binary.
    '
    'RETURN  : String = XML string containing, the set of Input
    '          Parameters or <ERROR> tag, for each binary PRM file.
    '          It may also return an <ERROR> tag for any general
    '          failure condition.
    '
    '            Sample Return XML structure:
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <PRM_PARAMS_LIST>
    '
    '                <!-- Successfully converted Lease PRM file -->
    '                <PRM_PARAMS>
    '                    <PRM_FILE_NAME>LeasePRMFile.prm</PRM_FILE_NAME>
    '                    <TransactionAmount>25000000</TransactionAmount>
    '                    <TransactionStartDate>2002-08-20</TransactionStartDate>
    '                    <ResidualAmount>100000</ResidualAmount>
    '                    <Structure>High/Low</Structure>
    '                    <Periodicity>Semiannual</Periodicity>
    '                    <PaymentTiming>Advance</PaymentTiming>
    '                    <NumberOfPayments>14</NumberOfPayments>
    '                    <TargetData>
    '                        <TypeOfStatistic>Yield</TypeOfStatistic>
    '                        <StatisticIndex>1</StatisticIndex>
    '                        <NEPA>Pre-tax nominal</NEPA>
    '                        <TargetValue>0.075</TargetValue>
    '                        <Adjust>Rent</Adjust>
    '                        <AdjustmentMethod>Proportional</AdjustmentMethod>
    '                    </TargetData>
    '                </PRM_PARAMS>
    '
    '                <!-- Successfully converted Loan PRM file -->
    '                <PRM_PARAMS>
    '                    <PRM_FILE_NAME>LoanPRMFile.prm</PRM_FILE_NAME>
    '                    <TransactionAmount>25000000</TransactionAmount>
    '                    <TransactionStartDate>2002-08-20</TransactionStartDate>
    '                    <ResidualAmount>100000</ResidualAmount>
    '                    <Structure>High/Low</Structure>
    '                    <Periodicity>Semiannual</Periodicity>
    '                    <PaymentTiming>Advance</PaymentTiming>
    '                    <NumberOfPayments>14</NumberOfPayments>
    '                    <TargetData>
    '                        <TypeOfStatistic>Yield</TypeOfStatistic>
    '                        <StatisticIndex>1</StatisticIndex>
    '                        <NEPA>Pre-tax nominal</NEPA>
    '                        <TargetValue>0.075</TargetValue>
    '                        <Adjust>Rent</Adjust>
    '                        <AdjustmentMethod>Proportional</AdjustmentMethod>
    '                    </TargetData>
    '                </PRM_PARAMS>
    '
    '                <!-- Error reading PRM file -->
    '                <PRM_PARAMS>
    '                    <PRM_FILE_NAME>ErrorPRMFile.prm</PRM_FILE_NAME>
    '                    <ERROR>
    '                        <ERROR_NBR>-1072896682</ERROR_NBR>
    '                        <ERROR_DESC>Error!!!...</ERROR_DESC>
    '                    </ERROR>
    '                </PRM_PARAMS>
    '                …
    '            </PRM_PARAMS_LIST>
    '
    '            OR In case of general failure:
    '
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <PRM_PARAMS_LIST>
    '                <ERROR>
    '                    <ERROR_NBR>-1072896682</ERROR_NBR>
    '                    <ERROR_DESC>Error!!!...</ERROR_DESC>
    '                </ERROR>
    '            </PRM_PARAMS_LIST>
    '================================================================
    <SoapDocumentMethod(Binding:="ISuperTrumpServiceSoapPort"), STAThreadAttribute(), WebMethod()> _
    Public Function GetPRMParams(ByVal astrPRMParamsInfoXML As String) As String


        Dim lobjSTWebService = ISuperTrumpService.Instance
        Try
            Return lobjSTWebService.GetPRMParams(astrPRMParamsInfoXML)
        Catch ex As Exception
            Return ex.Message()
        Finally
            lobjSTWebService = Nothing
        End Try

    End Function

    '================================================================
    'METHOD  : ModifyPRMFiles
    'PURPOSE : To modify the parameters contained in the binary PRM
    '          files and return the modified binary PRM file and/or
    '          XML equivalent for the binary PRM file and/or write
    '          to a file location.
    'PARMS   :
    '          astrModifyPRMFilesXML [String] = XML string containing
    '          the PRM file [either binary PRM file(s) or just path
    '          to the binary PRM file(s)], Parameters to be modified
    '          and the type of output that is expected (modified
    '          binary PRM file and/or XML equivalent for the binary PRM file and/or write to a file location).
    'RETURN  : String
    '================================================================
    <SoapDocumentMethod(Binding:="ISuperTrumpServiceSoapPort"), STAThreadAttribute(), WebMethod()> _
    Public Function ModifyPRMFiles(ByVal astrModifyPRMFilesXML As String) As String


        Dim lobjSTWebService = ISuperTrumpService.Instance
        Try
            Return lobjSTWebService.ModifyPRMFiles(astrModifyPRMFilesXML)
        Catch ex As Exception
            Return ex.Message()
        Finally
            lobjSTWebService = Nothing
        End Try

    End Function


    '================================================================
    'METHOD  : Test
    'PURPOSE : Returns a string that this component is able to invoke
    '          STServer (Ivory's SuperTrump Server component).
    'PARMS   : NONE
    'RETURN  : String
    '================================================================
    <SoapDocumentMethod(Binding:="ISuperTrumpServiceSoapPort"), STAThreadAttribute(), WebMethod()> _
    Public Function Test() As String

        Dim lobjSTWebService = ISuperTrumpService.Instance
        Try
            Return lobjSTWebService.Test()
        Catch ex As Exception
            Return ex.Message()
        Finally
            lobjSTWebService = Nothing
        End Try
    End Function

    '================================================================
    'METHOD  : RunAdHocXMLInOutQuery
    'PURPOSE : to allow adhoc XML queries to be submitted via the
    '           XMLInOut method in STServer. Values that are not
    '           available in the ConvertPRMToXML xml structure can
    '           be set/received this way. Files can be read,
    '           modified and new versions written to disk
    'PARMS   :
    '          astrXMLInOutQuery [String] = XML string containing the
    '          query to be executed and the file to be read
    '
    '          Sample Input Parameter structure:
    '            <PRM_INFO>
    '                <PRM_FILE>
    '                    <AD_HOC_QUERY>
    '                        <SuperTRUMP>
    '                            <Transaction id="TRAN4">
    '                                <ReadFile filename="\\ce213043914auct\Pricing$\test.prm"/>
    '                                <TransactionAmount query="true"/>
    '                            </Transaction>
    '                        </SuperTRUMP>
    '                    </AD_HOC_QUERY>
    '                </PRM_FILE>
    '            </PRM_INFO>
    '
    'RETURN  : String= XML string containing, the PRM query result or
    '          <ERROR> tag, for each set of PRM Input Parameters.
    '          It may also return an <ERROR> tag for any general
    '          failure condition.
    '================================================================
    <SoapDocumentMethod(Binding:="ISuperTrumpServiceSoapPort"), STAThreadAttribute(), WebMethod()> _
    Public Function RunAdHocXMLInOutQuery(ByVal astrXMLInOutQuery As String) As String


        Dim lobjSTWebService = ISuperTrumpService.Instance
        Try
            Return lobjSTWebService.RunAdHocXMLInOutQuery(astrXMLInOutQuery)
        Catch ex As Exception
            Return ex.Message()
        Finally
            lobjSTWebService = Nothing
        End Try

    End Function


    '================================================================
    'METHOD  : GeneratePRMFilesForPmtStruct
    'PURPOSE : To generate binary PRM file for each set of PRM
    '          parameters and Meta data. The PRM parameters contains
    '          the payment structure. This method is solving for
    '          payments.
    'PARMS   :
    '          astrPRMInfoXML [String] = XML string containing the
    '          PRM Parameters and Meta data required to generate the
    '          binary PRM file(s).
    '
    '          Sample Input Parameter structure:
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <PRM_INFO>
    '                <PRM_FILE>
    '                    <PRM_META_DATA>
    '                        <FILE_NAME>MyPRMFile.prm</FILE_NAME>
    '                        <TEMPLATE_NAME>USA 5 MACRS.TEM</TEMPLATE_NAME>
    '                        <MODE>Lessor</MODE>
    '                    </PRM_META_DATA>
    '                    <PRM_PARAMS>
    '                        <TRANSACTIONAMOUNT>25000000</TRANSACTIONAMOUNT>
    '                        <TRANSACTIONSTARTDATE>2002-08-20</TRANSACTIONSTARTDATE>
    '                        <PERIODICITY>Monthly</PERIODICITY>
    '                        <PAYMENTTIMING>Advance</PAYMENTTIMING>
    '                        <STRUCTURE>Level</STRUCTURE>
    '                        ...
    '                    </PRM_PARAMS>
    '                </PRM_FILE>
    '                <PRM_FILE>
    '                    <PRM_META_DATA>
    '                        <FILE_NAME>ErrorPRMFile.prm</FILE_NAME>
    '                        ...
    '                    </PRM_META_DATA>
    '                    …
    '                </PRM_FILE>
    '                …
    '            </PRM_INFO>
    'RETURN  : String= XML string containing, the binary PRM File or
    '          <ERROR> tag, for each set of PRM Input Parameters.
    '          It may also return an <ERROR> tag for any general
    '          failure condition.
    '
    '            Sample Return XML structure:
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <PRM_FILE_LIST>
    '
    '                <!-- Sucessful generation of PRM file -->
    '                <PRM_FILE>
    '                    <FILE_NAME>MyPRMFile.prm</FILE_NAME>
    '                    <FILE_DATA>/CQAGAAAAAAAAAAAAAAACAAAA3AAAAAAA…</FILE_DATA>
    '                </PRM_FILE>
    '
    '                <!-- Error generating PRM file -->
    '                <PRM_FILE>
    '                    <FILE_NAME>ErrorPRMFile.prm </FILE_NAME>
    '                    <ERROR>
    '                        <ERROR_NBR>-1072896682</ERROR_NBR>
    '                        <ERROR_DESC>Error!!!...</ERROR_DESC>
    '                    </ERROR>
    '                </PRM_FILE>
    '                …
    '            </PRM_FILE_LIST>
    '
    '            OR In case of general failure:
    '
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <PRM_FILE_LIST>
    '                <ERROR>
    '                    <ERROR_NBR>-1072896682</ERROR_NBR>
    '                    <ERROR_DESC>Error!!!...</ERROR_DESC>
    '                </ERROR>
    '            </PRM_FILE_LIST>
    '================================================================
    <SoapDocumentMethod(Binding:="ISuperTrumpServiceSoapPort"), STAThreadAttribute(), WebMethod()> _
   Public Function GeneratePRMFilesForPmtStruct(ByVal astrXMLInOutQuery As String) As String


        Dim lobjSTWebService = ISuperTrumpService.Instance
        Try
            Return lobjSTWebService.GeneratePRMFilesForPmtStruct(astrXMLInOutQuery)
        Catch ex As Exception
            Return ex.Message()
        Finally
            lobjSTWebService = Nothing
        End Try

    End Function

    '================================================================
    'METHOD  : GeneratePRMFilesForPmtStruct2
    'PURPOSE : To generate binary PRM file for each set of PRM
    '          parameters and Meta data. The PRM parameters contains
    '          the payment structure. This method is solving for
    '          rate.
    'PARMS   :
    '          astrPRMInfoXML [String] = XML string containing the
    '          PRM Parameters and Meta data required to generate the
    '          binary PRM file(s).
    '
    '          Sample Input Parameter structure:
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <PRM_INFO>
    '                <PRM_FILE>
    '                    <PRM_META_DATA>
    '                        <FILE_NAME>MyPRMFile.prm</FILE_NAME>
    '                        <TEMPLATE_NAME>USA 5 MACRS.TEM</TEMPLATE_NAME>
    '                        <MODE>Lessor</MODE>
    '                    </PRM_META_DATA>
    '                    <PRM_PARAMS>
    '                        <TRANSACTIONAMOUNT>25000000</TRANSACTIONAMOUNT>
    '                        <TRANSACTIONSTARTDATE>2002-08-20</TRANSACTIONSTARTDATE>
    '                        <PERIODICITY>Monthly</PERIODICITY>
    '                        <PAYMENTTIMING>Advance</PAYMENTTIMING>
    '                        <STRUCTURE>Level</STRUCTURE>
    '                        ...
    '                    </PRM_PARAMS>
    '                </PRM_FILE>
    '                <PRM_FILE>
    '                    <PRM_META_DATA>
    '                        <FILE_NAME>ErrorPRMFile.prm</FILE_NAME>
    '                        ...
    '                    </PRM_META_DATA>
    '                    …
    '                </PRM_FILE>
    '                …
    '            </PRM_INFO>
    'RETURN  : String= XML string containing, the binary PRM File or
    '          <ERROR> tag, for each set of PRM Input Parameters.
    '          It may also return an <ERROR> tag for any general
    '          failure condition.
    '
    '            Sample Return XML structure:
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <PRM_FILE_LIST>
    '
    '                <!-- Sucessful generation of PRM file -->
    '                <PRM_FILE>
    '                    <FILE_NAME>MyPRMFile.prm</FILE_NAME>
    '                    <FILE_DATA>/CQAGAAAAAAAAAAAAAAACAAAA3AAAAAAA…</FILE_DATA>
    '                </PRM_FILE>
    '
    '                <!-- Error generating PRM file -->
    '                <PRM_FILE>
    '                    <FILE_NAME>ErrorPRMFile.prm </FILE_NAME>
    '                    <ERROR>
    '                        <ERROR_NBR>-1072896682</ERROR_NBR>
    '                        <ERROR_DESC>Error!!!...</ERROR_DESC>
    '                    </ERROR>
    '                </PRM_FILE>
    '                …
    '            </PRM_FILE_LIST>
    '
    '            OR In case of general failure:
    '
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <PRM_FILE_LIST>
    '                <ERROR>
    '                    <ERROR_NBR>-1072896682</ERROR_NBR>
    '                    <ERROR_DESC>Error!!!...</ERROR_DESC>
    '                </ERROR>
    '            </PRM_FILE_LIST>
    '================================================================
    <SoapDocumentMethod(Binding:="ISuperTrumpServiceSoapPort"), STAThreadAttribute(), WebMethod()> _
 Public Function GeneratePRMFilesForPmtStruct2(ByVal astrXMLInOutQuery As String) As String


        Dim lobjSTWebService = ISuperTrumpService.Instance
        Try
            Return lobjSTWebService.GeneratePRMFilesForPmtStruct2(astrXMLInOutQuery)
        Catch ex As Exception
            Return ex.Message()
        Finally
            lobjSTWebService = Nothing
        End Try

    End Function
#End Region
End Class