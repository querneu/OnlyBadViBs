Set objUser = CreateObject("WScript.Network")
userName = objUser.UserName
domainName = objUser.UserDomain

FUNCTION GetUserDN(BYVAL UN, BYVAL DN)
Set ObjTrans = CreateObject("NameTranslate")
objTrans.init 1, DN
objTrans.set 3, DN & "\" & UN
strUserDN = objTrans.Get(1)
GetUserDN = strUserDN
END FUNCTION


Set objLDAPUser = GetObject("LDAP://" & GetUserDN(userName,domainName))

'Getting prepared to write the files
Dim objFSO, objWsh, appDataPath, pathToCopyTo, plainTextFile, plainTextFilePath, htmlFile, htmlFilePath
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWsh = CreateObject("WScript.Shell")
appDataPath = objWsh.ExpandEnvironmentStrings("%APPDATA%")
'pathToCopyTo = appDataPath & "\Microsoft\Signatures\"
pathToCopyTo = "C:\Signatures\"


'HTML signature � signature.htm
htmlFilePath = pathToCopyTo & "signature.htm"
Set htmlFile = objFSO.CreateTextFile(htmlFilePath, TRUE)
htmlfile.WriteLine("<body>")
'Nome e Cargo ou departamento
htmlfile.WriteLine("<p style='margin-top:5.0pt;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;margin-bottom:.0001pt'>"& vbNewLine & _ 
                        "<b>"& vbNewLine & _ 
                            "<span style='font-size:10.0pt;font-family:""Arial"",sans-serif;color:#003296'>" & objLDAPUser.DisplayName & "<br>"& vbNewLine & _ 
                            "</span>"& vbNewLine & _ 
                        "</b><span style='font-size:10.0pt;font-family:""Arial"",sans-serif;color:#868695'>" & objLDAPUser.title & "<o:p></o:p></span>")
'Sala Webex
htmlfile.WriteLine("<p style='margin-top:5.5pt;margin-right:0cm;margin-bottom:2.5pt;margin-left:0cm'>"& vbNewLine & _ 
                        "<span style='font-size:9.0pt;font-family:""Arial"",sans-serif;color:#868695'>"& vbNewLine & _ 
                            "<a href='https://sparkatelecom.webex.com/join/" & objLDAPUser.sAMAccountName & "' style='margin:none'>"& vbNewLine & _ 
                                "<span style='color:#868695;text-decoration:none;text-underline:none'>Sala Pessoal </span>"& vbNewLine & _ 
                                "<b>"& vbNewLine & _ 
                                    "<span style='color:#003296;text-decoration:none;text-underline:none'>Webex</span>"& vbNewLine & _ 
                                "</b>"& vbNewLine & _ 
                                "<span style='color:#868695;text-decoration:none;text-underline:none'> </span>"& vbNewLine & _ 
                            "</a>"& vbNewLine & _ 
                        "</span>"& vbNewLine & _ 
                    "</p>")
'Telefone Comercial e Site
htmlfile.WriteLine("<p style='margin-top:5.5pt;margin-right:0cm;margin-bottom:2.5pt;margin-left:0cm'>"& vbNewLine & _ 
                        "<b>"& vbNewLine & _ 
                            "<span style='font-size:9.0pt;font-family:""Arial"",sans-serif;color:#003296'>T:</span>"& vbNewLine & _ 
                        "</b>"& vbNewLine & _ 
                        "<span style='font-size:9.0pt;font-family:""Arial"",sans-serif;color:#868695'>" & objLDAPUser.telephoneNumber & "<br></span>"& vbNewLine & _ 
                        "<b>"& vbNewLine & _ 
                            "<span style='font-size:9.0pt;font-family:""Arial"",sans-serif;color:#003296'>W:</span>"& vbNewLine & _ 
                        "</b>"& vbNewLine & _ 
                        "<span style='font-size:9.0pt;font-family:""Arial"",sans-serif;color:#868695'> "& vbNewLine & _ 
                            "<a href=""http://www.atelecom.com.br"" title=""Visite nosso site""style='line-height:100%'>"& vbNewLine & _ 
                                "<span style='color:#868695;text-decoration:none;text-underline:none'>atelecom.com.br</span>"& vbNewLine & _ 
                            "</a> "& vbNewLine & _ 
                        "</span>"& vbNewLine & _ 
                    "</p>")
'Banner ATelecom
htmlfile.WriteLine("<table border=0 cellspacing=0 cellpadding=0 width=326 style='width:244.5pt;mso-cellspacing:0cm;background:white;mso-yfti-tbllook:1184;mso-padding-alt:0cm 0cm 0cm 0cm'>"& vbNewLine & _ 
                        "<tr>"& vbNewLine & _ 
                            "<td >"& vbNewLine & _ 
                                "<p class=MsoNormal>"& vbNewLine & _ 
                                    "<span>"& vbNewLine & _ 
                                        "<a href=""http://www.atelecom.com.br"">"& vbNewLine & _ 
                                            "<span>"& vbNewLine & _ 
                                                "<img border=0 width=326 height=86 id=""_x0000_i1028"" src=""http://www.atelecom.com.br/Assinaturas_A.Telecom/img-assinatura-email.jpg"" style='margin-bottom:0px;margin-left:0px;margin-right:0px;margin-top:0px' alt=""Banner ATelecom"">"& vbNewLine & _ 
                                            "</span>"& vbNewLine & _ 
                                        "</a>"& vbNewLine & _ 
                                    "</span>"& vbNewLine & _ 
                                "</p>"& vbNewLine & _ 
                            "</td>"& vbNewLine & _ 
                        "</tr>"& vbNewLine & _ 
                        "<tr class='MsoNormalTable'>"& vbNewLine & _ 
                            "<td>"& vbNewLine & _ 
                                "<p class=MsoNormal>"& vbNewLine & _ 
                                    "<span style='mso-fareast-font-family:""Times New Roman""'>"& vbNewLine & _ 
                                        "<a href=""https://www.facebook.com/atelecom"">"& vbNewLine & _ 
                                            "<span style='text-decoration:none;text-underline:none'>"& vbNewLine & _ 
                                                "<img border=0 width=30 height=30 id=""_x0000_i1029"" src=""http://www.atelecom.com.br/Assinaturas_A.Telecom/img-assinatura-email-fb-exp.png"" style='margin-bottom:0px;margin-left:0px;margin-right:0px;margin-top:0px' alt=""Facebook icon""></span>"& vbNewLine & _ 
                                        "</a>"& vbNewLine & _ 
                                    "</span>"& vbNewLine & _ 
                                "</p>"& vbNewLine & _ 
                            "</td>"& vbNewLine & _ 
                            "<td>"& vbNewLine & _ 
                                "<p class=MsoNormal>"& vbNewLine & _ 
                                    "<span style='mso-fareast-font-family:""Times New Roman""'>"& vbNewLine & _ 
                                        "<a href=""https://www.linkedin.com/company/a.telecom/"">"& vbNewLine & _ 
                                            "<span style='text-decoration:none;text-underline:none'>"& vbNewLine & _ 
                                                "<img border=0 width=30 height=30 id=""_x0000_i1030"" src=""http://www.atelecom.com.br/Assinaturas_A.Telecom/img-assinatura-email-linkedin-exp.png"" style='margin-bottom:0px;margin-left:0px;margin-right:0px;margin-top:0px' alt=""Linkedin Icon"">"& vbNewLine & _ 
                                            "</span>"& vbNewLine & _ 
                                        "</a>"& vbNewLine & _ 
                                    "</span>"& vbNewLine & _ 
                                "</p>"& vbNewLine & _ 
                            "</td>"& vbNewLine & _ 
                            "<td>"& vbNewLine & _ 
                                "<p class=MsoNormal>"& vbNewLine & _ 
                                    "<span style='mso-fareast-font-family:""Times New Roman""'>"& vbNewLine & _ 
                                        "<a href=""https://www.instagram.com/atelecom_"">"& vbNewLine & _ 
                                            "<span>"& vbNewLine & _ 
                                                "<img border=0 width=30 height=30 id=""_x0000_i1030"" src=""http://www.atelecom.com.br/Assinaturas_A.Telecom/img-assinatura-email-instagram-exp.jpg"" style='margin-bottom:0px;margin-left:0px;margin-right:0px;margin-top:0px' alt=""Linkedin Icon"">"& vbNewLine & _ 
                                            "</span>"& vbNewLine & _ 
                                        "</a>"& vbNewLine & _ 
                                    "</span>"& vbNewLine & _ 
                                "</p>"& vbNewLine & _ 
                            "</td>"& vbNewLine & _ 
                        "</div>"& vbNewLine & _ 
                    "</table>")
'htmlfile.WriteLine("<div><strong>" & objLDAPUser.description & "</strong></div>")
'htmlfile.WriteLine("<div><strong>" & objLDAPUser.ipPhone & "</strong></div>")
'htmlfile.WriteLine("<div><strong>" & objLDAPUser.mobile & "</strong></div>")
'htmlfile.WriteLine("<div><strong>" & objLDAPUser.department & "</strong></div>")
'htmlfile.WriteLine("<div><strong>" & objLDAPUser.title & "</strong></div>")
htmlfile.WriteLine("</body>")
htmlfile.WriteLine("</html>")

'RTF signature � Copies over pre-made RTF signature
'Set fso = CreateObject("Scripting.FileSystemObject")
'fso.CopyFile "\\fileserver\share\signature.rtf", appDataPath & "\Microsoft\Signatures\",TRUE
