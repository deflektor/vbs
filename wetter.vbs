Dim WshShell

Set WshShell = createobject("wscript.shell") 

Set xmlDoc = CreateObject("Msxml2.DOMDocument") 
xmlDoc.async = False
xmlDoc.setProperty "ServerHTTPRequest", True
xmlDoc.Load "http://xml.weather.yahoo.com/forecastrss?p=AUXX0025&u=c"
'xmlDoc.setProperty "SelectionLanguage", "XPath"

Set ElemList = xmlDoc.getElementsByTagName("yweather:forecast")
wscript.echo ("Found " + ElemList.item(0).getAttribute("low"))
wscript.echo ("Found " + ElemList.item(0).getAttribute("high"))
wscript.echo ("Found " + ElemList.item(0).getAttribute("text"))