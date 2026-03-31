# XmlHttpRequest – Http requests in Excel VBA (Updated 2022)

[Source](https://codingislove.com/http-requests-excel-vba/)

Excel is a powerful and most popular tool for data analysis! HTTP requests in VBA gives additional capabilities to Excel. XmlHttpRequest object is used to make HTTP requests in VBA. HTTP requests can be used to interact with a web service, API or even websites. Let’s understand how it works.

Open an excel file and open VBA editor (`Alt + f11`) > `New Module` and start writing code in a `sub`

```vb
Public sub XmlHttpTutorial

End Sub
Define XMLHttpRequest
Define http client using following code

Dim xmlhttp as object
Set xmlhttp = CreateObject("MSXML2.serverXMLHTTP")
```

> If you need VBA’s Intellisense autocomplete then do it this way :
> 
> 1. Add a reference to MSXML (`Tools` > `references`)
> 
> 2. Select appropriate version based on your PC :
>   - Microsoft XML, v 3.0.
>   - Microsoft XML, v 4.0 (if you have installed MSXML 4.0 separately).
>   - Microsoft XML, v 5.0 (if you have installed Office 2003 – 2007 which provides MSXML 5.0 for Microsoft Office Applications).
>   - Microsoft XML, v 6.0 for latest versions of MS Office.

Then **define http client**:

```vb
Dim xmlhttp As New MSXML2.XMLHTTP
'Dim xmlhttp As New MSXML2.XMLHTTP60 for Microsoft XML, v 6.0 
```

VBA Intellisense will show you the right one when you start typing.

**Make requests**

Requests can be made using open and send methods. Open method syntax is as follows :

```vb
xmlhttp.Open Method, URL, async(true or false)
```
I’m using requestBin to test requests. Create a bin there and send requests to that URL to test requests.

## GET

A simple **GET** request would be :

```vb
Dim xmlhttp As New MSXML2.XMLHTTP60, myurl As String
myurl = "http://requestb.in/15oxrjh1" //replace with your URL
xmlhttp.Open "GET", myurl, False
xmlhttp.Send
MsgBox(xmlhttp.responseText)
```

Run this code, a message box is displayed with the response of the request.

## Headers

**Request headers**

Request headers can be set using setRequestHeader method. Examples :

```vb
xmlhttp.setRequestHeader "Content-Type", "text/json"
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0 (iPad; U; CPU OS 3_2_1 like Mac OS X; en-us) AppleWebKit/531.21.10 (KHTML, like Gecko) Mobile/7B405"
xmlhttp.setRequestHeader "Authorization", AuthCredentials
```
## POST

Simple **POST** request to send **formdata**: POST requests are used to send some data, data can be sent in Send method. A simple POST request to send form data :

```vb
Public Sub httpclient()
Dim xmlhttp As New MSXML2.XMLHTTP, myurl As String
myurl = "http://requestb.in/15oxrjh1"
xmlhttp.Open "POST", myurl, False
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xmlhttp.Send "name=codingislove&email=admin@codingislove.com"
MsgBox (xmlhttp.responseText)
End Sub
```

## Authentication

**Basic Authentication in VBA**. When we need to access web services with basic authentication, A username and password have to be sent with the Authorization header. Username and password should also be base64 encoded. Example :

```vb
user = "someusername"
password = "somepassword"
xmlhttp.setRequestHeader "Authorization", "Basic " + Base64Encode(user + ":" + password)
```

# Practical Use Cases

Practical use cases of http requests in VBA are unlimited. Some of them are pulling data from Yahoo finance API, weather API, pulling orders from Ecommerce store admin panel, uploading products, retrieving web form data to excel etc.

- [Parse HTML in Excel VBA](https://codingislove.com/parse-html-in-excel-vba/) – Learn by parsing hacker news home page where I retrieve a web page using HTTP GET request and parse its HTML to get data from a web page.

- [How to build a simple weather app in Excel VBA](https://codingislove.com/weather-app-in-excel-vba/) where I make a HTTP Get request to weather API

- [JSON Api in Excel VBA](https://codingislove.com/excel-json/) where I call JSON Apis using HTTP GET and POST requests.

---

[DOC MOC](./doc-00_MOC.md)