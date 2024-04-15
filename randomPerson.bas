Attribute VB_Name = "randomPerson"
Function randomFirstName()

    Dim HTTPreq As Object, url As String, response As String, nameType As String
    nameType = "firstname"
    
    'Setting API parameters
    Set HTTPreq = CreateObject("MSXML2.XMLHTTP")
    url = "https://randommer.io/api/Name?nameType=" & nameType & "&quantity=1&time=" & Timer
    response = ""
    
    With HTTPreq
        .Open "GET", url, False
        .SetRequestHeader "X-Api-Key", "ae3b7ad5958740498bc2f4f79dba51a3"
        
        .Send
    End With
    
    'Api's response
    response = HTTPreq.ResponseText
    
    'Delete unnecessery characters from the response
    response = Replace(response, Chr(91) & Chr(34), "")
    response = Replace(response, Chr(34) & Chr(93), "")
    'Return reponse for the function
    randomFirstName = response
    
End Function

Function randomSurName()

    Dim HTTPreq As Object, url As String, response As String, nameType As String
    nameType = "surname"
    
    'Setting API parameters
    Set HTTPreq = CreateObject("MSXML2.XMLHTTP")
    url = "https://randommer.io/api/Name?nameType=" & nameType & "&quantity=1&time=" & Timer
    response = ""
    
    With HTTPreq
        .Open "GET", url, False
        .SetRequestHeader "X-Api-Key", "ae3b7ad5958740498bc2f4f79dba51a3"
        
        .Send
    End With
    
    'Api's response
    response = HTTPreq.ResponseText
    
    'Delete unnecessery characters from the response
    response = Replace(response, Chr(91) & Chr(34), "")
    response = Replace(response, Chr(34) & Chr(93), "")
    'Return reponse for the function
    randomSurName = response
    
End Function
