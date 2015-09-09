Dim speaks, speech, sp_date, sp_weather
speaks="Welcome, Master"
Set speech=CreateObject("sapi.spvoice")
speech.Speak speaks
sp_date = "Today is the " + FormatDateTime(Date(), 2)
speech.Speak sp_date

Set xmlDoc = CreateObject("Msxml2.DOMDocument") 
xmlDoc.async = False
xmlDoc.setProperty "ServerHTTPRequest", True
xmlDoc.Load "http://xml.weather.yahoo.com/forecastrss?p=AUXX0025&u=c"
'xmlDoc.setProperty "SelectionLanguage", "XPath"

Set ElemList = xmlDoc.getElementsByTagName("yweather:forecast")
sp_weather = "Today weather is " + ElemList.item(0).getAttribute("text")
speech.Speak sp_weather
sp_weather = " Todays low temperature " + ElemList.item(0).getAttribute("low") + " degrees "
speech.Speak sp_weather
sp_weather = " Todays high temperature " + ElemList.item(0).getAttribute("high") + " degrees "
speech.Speak sp_weather

sp_date = "Have a? nice Day!"
speech.Speak sp_date