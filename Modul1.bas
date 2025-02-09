Attribute VB_Name = "Modul11"
Function GPT(query As String, Optional apiKey As String = "", Optional apiUrl As String = "http://127.0.0.1:1234/v1/chat/completions", Optional model As String = "phi-4", Optional useJsonFormat As Boolean = False, Optional jsonKey As String = "") As String
    Dim httpObject As Object
    Set httpObject = CreateObject("MSXML2.ServerXMLHTTP")

    ' Encode query to handle special characters and line breaks
    query = Replace(query, vbCrLf, "\n")
    query = Replace(query, vbCr, "\n")
    query = Replace(query, vbLf, "\n")
    query = Replace(query, """", "\""")
    
    Dim requestBody As String
    If useJsonFormat Then
        requestBody = "{""model"": """ & model & """, ""messages"": [{""role"": ""user"", ""content"": """ & query & """}], ""response_format"": {""type"": ""json_object""}}"
    Else
        requestBody = "{""model"": """ & model & """, ""messages"": [{""role"": ""user"", ""content"": """ & query & """}]}"
    End If
    
    ' Send a POST request
    httpObject.Open "POST", apiUrl, False
    httpObject.setRequestHeader "Content-Type", "application/json"
    If apiKey <> "" Then
        httpObject.setRequestHeader "Authorization", "Bearer " & apiKey
    End If
    
    On Error GoTo ErrorHandler
    httpObject.send (requestBody)
    
    Dim response As String
    response = httpObject.responseText
    
    ' Extract the response (you may need to adjust the parsing based on the response structure)
    Dim json As Object
    Set json = JsonConverter.ParseJson(response)
    
    If useJsonFormat Then
        If jsonKey <> "" Then
            If Not json Is Nothing And IsObject(json) Then
                On Error Resume Next
                GPT = json(jsonKey)
                If Err.Number <> 0 Then
                    GPT = "Error: Specified key not found in JSON."
                End If
                On Error GoTo 0
            Else
                GPT = "Error: Invalid JSON response."
            End If
        Else
            GPT = response
        End If
    Else
        If Not json Is Nothing And IsObject(json) Then
            On Error Resume Next
            GPT = json("choices")(1)("message")("content")
            If Err.Number <> 0 Then
                GPT = "Error: Unable to parse response content."
            End If
            On Error GoTo 0
        Else
            GPT = "Error: Invalid JSON response."
        End If
    End If
    
    Set httpObject = Nothing
    Exit Function
    
ErrorHandler:
    GPT = "Error: " & Err.Description
    Set httpObject = Nothing
End Function
