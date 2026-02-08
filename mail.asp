<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="utf-8">
<title>Classic Asp Mail Gönderme</title>
<meta name="description" content="Classic Asp Mail Gönderme" />
<meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=yes">
</head>


<%
islem = request.querystring("islem")

if islem ="mail" then

    if request.Form("gv") <> request.Form("gv1") then
    %><script>
    alert('Güvenlik Kodunu Yanlış Girdiniz...')
    history.back()
    </script><% 
    response.end
    end if	
    
        if request.form("isim")= "" or  request.form("mail")= "" then %>
        <script> alert('Lütfen Boş Alanları Doldurunuz...')
        history.back()</script> <%
        response.end
        end if

aisim= request.form("isim")
amail= request.form("mail")
atel= request.form("tel")
akonu= request.form("konu")

body= "<br>-- www.sunucuadresiniz.com --<br><br><b>Isim :</b> " & aisim & "<b><br>Mail :</b> " & amail & "<b><br>Tel :</b> " & atel & "<b><br>Konu :</b> " & akonu & "<b><br>Tarih : </b>" & date & "<b><br>Saat : </b>" & time &"<br><b>İp:</b> " & Request.ServerVariables("REMOTE_ADDR")	 

Dim iMsg, iConf, Flds
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields

schema = "http://schemas.microsoft.com/cdo/configuration/"
Flds.Item(schema & "sendusing") = 2
Flds.Item(schema & "smtpserver") = "mail.sunucuadresiniz.com" 
Flds.Item(schema & "smtpserverport") = 587
Flds.Item(schema & "smtpauthenticate") = 1
Flds.Item(schema & "sendusername") = "otomatik@sunucuadresiniz.com"
Flds.Item(schema & "sendpassword") =  "mailsifreniz"
Flds.Item(schema & "smtpusessl") = 0
Flds.Update
With iMsg
.To = "mailgonderilicekadres" 
.From = "otomatik@sunucuadresiniz.com"
.Subject = "mail konusu"
.HTMLBody = body
.Sender = "otomatik@sunucuadresiniz.com"
.Organization = ""
.ReplyTo = amail
.Server = "mail.sunucuadresiniz.com" 
.Username = "otomatik@sunucuadresiniz.com"
.Password =  "mailsifreniz"
Set .Configuration = iConf
SendEmailGmail = .Send
End With
set iMsg = nothing
set iConf = nothing
set Flds = nothing
  %>
  
Mesajınız Gönderildi
<%
else
%>	


<form method="POST" action="?islem=mail">
						<div class="form">
							<div class="col-md-6 noleftmargin">
								<label>İsim Soyisim</label>
								<input type="text" name="isim" required class="validate[required]">
							</div>
							<div class="col-md-6 noleftmargin">
								<label>E-Mail</label>
								<input type="text" name="mail" required class="validate[email,required]">
							</div>
							<div class="col-md-6 noleftmargin">
								<label>Telefon</label>
								<input type="text" name="tel">
							</div>
							<div class="col-md-6 noleftmargin">
								<label>Konu</label>
								<input type="text" name="konu">
							</div>
        	<div class="col-md-6 noleftmargin">
	        <label>Güvenlik Kodu</label>	
                              <% guvenlik=hexValue(5)
                              if guvenlik="" then
                              guvenlik="37ZSHF"
                              end if
                              %>
          <input type="text" name="gv"  required>
          <input type="hidden" name="gv1" size="10" value="<%=guvenlik%>" >
        </div>
	    <div class="col-md-6 noleftmargin">
	    <label> &nbsp;</label>
	    <%=guvenlik%>	
	    </div>
</form>

<%
end if
%>	

