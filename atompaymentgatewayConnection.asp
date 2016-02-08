<html>
<body>
Connection to Atom payment gateway........
<br/>
<br/>
<%
Dim data, httpRequest, postResponse, currentXML

data="login=7&pass=Test@123&ttype=NBFundTransfer&prodid=NSE&amt=100&txncurr=INR&txnscamt=0&clientcode=123&txnid=123456&date=29/10/2014 12:09:06&custacc=123456789012&mdd=123"

set httpRequest=Server.CreateObject("MSXML2.ServerXMLHTTP")
httpRequest.Open "POST", "http://192.168.150.46/paynetz/epi/fts", False
httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
httpRequest.Send data

postResponse=httpRequest.ResponseText

'Response.Write postResponse 

'----xml Parsing--------
set currentXML=Server.CreateObject("MSXML2.DOMDocument.3.0")
currentXML.LoadXml(postResponse)
Dim tempTxnId, token, txnDetails
'-------------------------------------
tempTxnId = currentXML.DocumentElement.ChildNodes(0).ChildNodes(0).ChildNodes(2).text
token = currentXML.DocumentElement.ChildNodes(0).ChildNodes(0).ChildNodes(3).text
txnDetails="ttype=NBFundTransfer&txnStage=1&tempTxnId=" & tempTxnId & "&token=" & token

response.write(txnDetails)
Response.redirect("http://192.168.150.46/paynetz/epi/fts?" & txnDetails)

%>
</body>
</html>