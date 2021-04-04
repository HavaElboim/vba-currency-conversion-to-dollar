Rem import phase 6. convert donation value in dvalue to USD equivalent in dval_usd

Rem **********************************************
' For this script to work, need to
' 1
' Import JsonConverter.bas and Dictionary.cls(Open VBA Editor,  File > Import File)
' (these libraries are available at:
'  https://github.com/VBA-tools/VBA-Dictionary  and
'  https://github.com/VBA-tools/VBA-JSON
' - download & unzip the latest version of each
' 2.
' Include a reference to "Microsoft Scripting Runtime" (via vba menu: Tools -> References)
' 3.
' To work with XML, need to add in reference to the XML library via:
' Tools > References to Microsoft XML v3.0

' 4. After importing the above references, close Access and reopen
Rem **********************************************

Private Sub Command16_Click()

' Bank of Israel exchange rates: are only from foreign currencies to shekel
' - they do not include rates from other currencies to USD

Rem see explanation here about extracting data from online xml file:
Rem https://chandoo.org/forum/threads/reading-xml-file-online-using-vba-getting-this-error.29772/

Rem and see explanation here of online exchange rates at Bank of Israel site:
Rem https://www.boi.org.il/he/Markets/Pages/explainxml.aspx
Rem so we can display online dollar-shekel rates for a given date, e.g. 23/12/2020 using:
Rem http://www.boi.org.il/currency.xml?rdate=20201223&curr=01

Rem the data comes from the Bank of Israel site in XML form:
' the xml file containing currencies will be in this format:
' <?xml version="1.0" encoding="utf-8" standalone="yes"?>
' <CURRENCIES>
'  <LAST_UPDATE>2020-12-23</LAST_UPDATE>
'  <CURRENCY>
'    <NAME>Dollar</NAME>
'    <UNIT>1</UNIT>
'    <CURRENCYCODE>USD</CURRENCYCODE>
'    <COUNTRY>USA</COUNTRY>
'    <RATE>3.222</RATE>
'    <CHANGE>-0.371</CHANGE>
'  </CURRENCY>
' </CURRENCIES>

'see here for another example of extracting similar data into access by vba:
' https://stackoverflow.com/questions/18432430/how-can-i-select-an-xml-element-by-attribute-without-iterating-using-php
       ' or see here:
    ' https://stackoverflow.com/questions/21580868/using-vba-to-import-xml-website-into-access
    Dim conversionLog As String

    Dim node As Object
    Dim year, month, day As String

    Dim emailX As String
    Dim last_emailx As String
    Dim currencyX As String
    Dim donationValueX As Single
    Dim donationDateX As String
    Dim conversionString1, conversionString2, conversionString3, conversionString2USD As String
    'conversionString1: first part of URL for currency exchange data
    'conversionString2USD: last part of URL for USD-NIS exchange data
    'conversionString2Other: last part of URL for Other-NIS exchange data (Other = GBP / Swiss Franc / other)
    
    'Variables for retrieval of rates data from online API:
    Dim URL As String
    ' For access to work with XML, need to add in reference to the XML library via:
    ' Tools > References)
    ' C:\windows\system32\msxml6.dll
    Dim xmlhttp As New MSXML2.xmlhttp
    Dim xmlDoc As MSXML2.DOMDocument
    Dim XMLNodeList As MSXML2.IXMLDOMNodeList
    Dim XMLNodeList2 As MSXML2.IXMLDOMNodeList
    Dim xmlElement As MSXML2.IXMLDOMElement
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' more variables needed for working with the XML files:
    Dim i As Long
    Dim strUrl As String   'this will be the link which will give the exchange rate from the Bank of Israel
                            ' for a given date and currency
    Dim conversionRateUSD As Single
                            
    ' Exchange rates from other currencies to dollar are available from the European Central Bank
    ' as an api at https://api.exchangeratesapi.io/
' see here for explanation: https://exchangeratesapi.io/
' to import from api to vba need library from https://github.com/VBA-tools/VBA-JSON :
' Import JsonConverter.bas into your project (Open VBA Editor, Alt + F11; File > Import File)
' and
' Add Dictionary reference/class
'   For Windows-only, include a reference to "Microsoft Scripting Runtime" (via vba menu: Tools -> References)
'   For Windows and Mac, include VBA-Dictionary: from https://github.com/VBA-tools/VBA-Dictionary - download & unzip & then:
'              ->  import Dictionary.cls into your VBA project.
' see more info on import from json here: https://www.youtube.com/watch?v=CFFLRmHsEAs&feature=youtu.be
                        
    ' variables for working with API for other currencies:
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    ' the api offers historical exchange rates at https://api.exchangeratesapi.io/history
    ' parameters:
    ' base=USD gives exchange rates against the US dollar
    ' symbols=GBP gives exchange rates for the British Pound
    ' start_at=2018-01-01&end_at2018-01-01 gives exchange rate for date 1/1/2018
    conversionString2 = "https://api.exchangeratesapi.io/"
    Dim Json As Object
   
    
    ' variables for working with the database:
    Dim id As Integer
    Dim d As dao.Database
    Dim t As dao.Recordset

    Set d = CurrentDb
    Set t = d.OpenRecordset("filter new donors")
 
    conversionString1 = "https://www.boi.org.il/currency.xml?rdate="
    
    If t.EOF Then GoTo skip
    t.MoveFirst
    'Debug.Print "Looping through donations"
    conversionLog = "Converting new donations from dvalue to equivalent USD value in dval_usd"
    
    Do While Not t.EOF
        
    currencyX = t!Currency

    If currencyX = "USD" Then
        'dval_usd will just be the same as dval
        t.Edit
        t!dval_usd = t!dvalue
        Debug.Print vbNewLine & "donation currency: USD, new dval_usd: " & t!dval_usd
        conversionLog = conversionLog & vbNewLine & "donation currency: USD, new dval_usd: " & t!dval_usd
      
        t.Update
               
    ElseIf currencyX = "ILS" Then
            conversionString2USD = "&curr=01"
            Debug.Print vbNewLine & "converting donation of " & t!dvalue & " from ILS to USD"
            conversionLog = conversionLog & vbNewLine & "converting donation of " & t!dvalue & " ILS to USD"
       ' to convert from NIS to USD
    
    ' Bank of Israel site gives currency rates TO SHEKEL ONLY
    ' So this only helps us to convert shekel donations to USD.
    ' currency conversion data in the BOI site is provided for currencies with the following codes
    ' to insert in the URL giving the data:
    '01      ãåìø?   àøöåú äáøéú?
    '(other currency codes irrelevant)
    
                donationValueX = t!dvalue
                donationDateX = Format(t![Donation Date], "yyyyMMdd")
NewDate:        ' DON'T REMOVE THIS LINE NewDate: - the script returns here is there is no conversion data for the previous date
                strUrl = conversionString1 & donationDateX & conversionString2USD
                Debug.Print "Fetching rate from Bank of Israel data at: " & strUrl
                conversionLog = conversionLog & vbNewLine & "Fetching rate from Bank of Israel data at: " & strUrl
                
                ' Fetch the XML
                xmlhttp.Open "Get", strUrl, False
                xmlhttp.Send
                
                'Debug.Print "Fetched XML tree. Response: " & xmlhttp.ResponseText
                
                ' XMLNodes basename: xml
                 'Debug.Print "XMLNodes basename: " & xmlhttp.responseXML.childNodes.nextNode.baseName
                 ' XMLNodes basename: CURRENCIES
                'Debug.Print "XMLNodes basename: " & xmlhttp.responseXML.childNodes.nextNode.nextSibling.baseName  'CURRENCIES
                'Debug.Print "Descending another level in the XML tree: "
                'Debug.Print "XMLNodes length: " & xmlhttp.responseXML.childNodes.nextNode.nextSibling.childNodes.length  '2
                 Set XMLNodeList = xmlhttp.responseXML.childNodes.nextNode.nextSibling.childNodes
                 'Debug.Print "number of nodes: " & XMLNodeList.length  '2
                 
                 If (XMLNodeList.length = 4) Then
                    'this will happen if there was no rate set at the requested date, try one day earlier
                    Debug.Print "No exchange rate was set for this currency on " & donationDateX
                    conversionLog = conversionLog & vbNewLine & "No exchange rate was set for this currency on " & donationDateX
                    donationDateX = donationDateX - 1
                    Debug.Print "Trying instead date " & donationDateX
                    conversionLog = conversionLog & vbNewLine & "Trying instead date " & donationDateX
                    GoTo NewDate
                End If
                'returns "CURRENCY":
                'Debug.Print "XMLNodes basename: " & XMLNodeList.nextNode.nextSibling.baseName   'CURRENCY
                
                Set XMLNodeList2 = XMLNodeList.nextNode.nextSibling.childNodes
                'Get the data
                ' -  the conversion rate should be stored in the Rate node in the xml file.
                i = 1
                For Each xmlElement In XMLNodeList2
                    'Debug.Print i, xmlElement.nodeName, xmlElement.nodeTypedValue
                    If xmlElement.nodeName = "RATE" Then
                        'Debug.Print "Conversion rate to NIS is: " & xmlElement.nodeTypedValue
                        conversionRate = xmlElement.nodeTypedValue
                    End If
                    i = i + 1
                Next xmlElement
            
                t.Edit
                t!dval_usd = donationValueX / conversionRate
 
               t.Update
               Debug.Print "Donation value " & donationValueX & "ILS. Rate for " & currencyX & " on " & donationDateX & " was: " & conversionRate & " Updated equivalent $ value dval_usd to " & t!dval_usd
                conversionLog = conversionLog & vbNewLine & "Donation value " & donationValueX & "NIS. Rate for " & currencyX & " on " & donationDateX & " was: " & conversionRate & " Updated equivalent $ value dval_usd to " & t!dval_usd
           
    Else
                ' Donation was in some other currency (not NIS or USD0
                conversionString3 = "?base=USD&symbols=" & currencyX
            'put together URL that accesses the data for the required currency on the given date:
                donationValueX = t!dvalue
                donationDateX = Format(t![Donation Date], "yyyy-MM-dd")
          ' The API automatically moves to an earlier date if there is no conversion data for the chosen date
                strUrl = conversionString2 & donationDateX & conversionString3
                Debug.Print vbNewLine & "Converting donation value from " & currencyX & " to USD. Fetching rate from European Central Bank  data at: " & strUrl
                conversionLog = conversionLog & vbNewLine & "Converting donation value from " & currencyX & " to USD. Fetching rate from European Central Bank  data at: " & strUrl
                
               MyRequest.Open "GET", strUrl

                MyRequest.Send
                'Debug.Print "conversion rate data: " & MyRequest.ResponseText
                Set Json = JsonConverter.ParseJson(MyRequest.ResponseText)
                    
                conversionRate = Json("rates")(currencyX)
            
                t.Edit
                t!dval_usd = donationValueX / conversionRate
                t.Update
                Debug.Print "Donation value " & donationValueX & ". Rate for " & currencyX & " on " & donationDateX & " was: " & conversionRate & " Updated equivalent $ value dval_usd to " & t!dval_usd
                conversionLog = conversionLog & vbNewLine & "Donation value " & donationValueX & currencyX & ". Rate for " & currencyX & " on " & donationDateX & " was: " & conversionRate & " Updated equivalent $ value dval_usd to " & t!dval_usd

        'Else 'no matching currency
        '    GoTo skip
        End If
        
            
    t.MoveNext
    Loop
    
skip:
  'break out of function
                
        MsgBox conversionLog
        t.Close
        
End Sub