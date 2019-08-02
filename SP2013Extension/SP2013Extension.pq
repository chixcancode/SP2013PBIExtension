// This file contains your Data Connector logic
section SP2013Extension;

[DataSource.Kind="SP2013Extension", Publish="SP2013Extension.Publish"]
shared SP2013Extension.Contents = (spurl as text, queryText as text) =>
      let
          searchUrl = Text.Combine({spurl, "/_api/search/query?querytext=", queryText}),
           AuthenticateUser = SP2013Extension.AuthenticateUser(spurl),
           SearchResults = Xml.Tables(Web.Contents(searchUrl, [Headers=[Cookie=AuthenticateUser]]) )
//            ErrorCode = AuthenticateUser[ErrorCode],
//            SearchResults = if(ErrorCode = "NoError") then Xml.Tables(Web.Contents(searchUrl)) else Xml.Tables(Web.Contents(searchUrl))
     in
       SearchResults;
shared SP2013Extension.AuthenticateUser = (spurl as text) =>
      let
           authenticationUrl = Text.Combine({spurl, "/_vti_bin/Authentication.asmx"}),
           username = Extension.CurrentCredential()[Username],
           password = Extension.CurrentCredential()[Password],
       
       
       soapContent = Text.Combine({"<?xml version=""1.0"" encoding=""utf-8""?>
                <soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">
                  <soap:Body>
                    <Login xmlns=""http://schemas.microsoft.com/sharepoint/soap/"">
                      <username>", username, "</username>
                      <password>", password, "</password>
                    </Login>
                  </soap:Body>
                </soap:Envelope>"}),
         Source = Value.Metadata(Web.Contents(authenticationUrl,[Content=Text.ToBinary(soapContent),Headers=[ SOAPAction="http://schemas.microsoft.com/sharepoint/soap/Login", #"Content-Type"="text/xml;charset=UTF-8"]])),
       ReponseHeaders = Source[Headers],
       Cookie = ReponseHeaders[#"Set-Cookie"]
     
    in 
        Cookie;
// Data Source Kind description
SP2013Extension = [
    Authentication = [
        UsernamePassword = [],
		Windows = []
    ],
    Label = Extension.LoadString("DataSourceLabel")
];

// Data Source UI publishing description
SP2013Extension.Publish = [
    Beta = true,
    Category = "Other",
    ButtonText = { Extension.LoadString("ButtonTitle"), Extension.LoadString("ButtonHelp") },
    LearnMoreUrl = "https://powerbi.microsoft.com/",
    SourceImage = SP2013Extension.Icons,
    SourceTypeImage = SP2013Extension.Icons
];

SP2013Extension.Icons = [
    Icon16 = { Extension.Contents("SP2013Extension16.png"), Extension.Contents("SP2013Extension20.png"), Extension.Contents("SP2013Extension24.png"), Extension.Contents("SP2013Extension32.png") },
    Icon32 = { Extension.Contents("SP2013Extension32.png"), Extension.Contents("SP2013Extension40.png"), Extension.Contents("SP2013Extension48.png"), Extension.Contents("SP2013Extension64.png") }
];
