on error resume next
 
Const schema   = "http://schemas.microsoft.com/cdo/configuration/"
Const cdoBasic = 1
Const cdoSendUsingPort = 2
Dim oMsg, oConf
 
' Propriedades do email
Set oMsg      = CreateObject("CDO.Message")
oMsg.From     = "douglasproglimateste@gmail.com" ' ou "Nome do remetente <from@gmail.com>"
oMsg.To       = "douglasproglima@gmail.com"      ' ou "Nome do destino <to@gmail.com>"
oMsg.Subject  = "Teste de VBscript"
oMsg.TextBody = "Mensagem enviada com sucesso !!! Enjoy it"
 
' Configuração e autenticação do seu servidor de SMTP Gmail
Set oConf = oMsg.Configuration
 
'Endereço do servidor de SMTP
oConf.Fields(schema & "smtpserver")       = "smtp.gmail.com"
 
'Número da porta
oConf.Fields(schema & "smtpserverport")   = 465
 
oConf.Fields(schema & "sendusing")        = cdoSendUsingPort
 
'Tipo de autenticacao
oConf.Fields(schema & "smtpauthenticate") = cdoBasic
 
'Uso da Encriptação SSL
oConf.Fields(schema & "smtpusessl")       = True
 
'Envia username
oConf.Fields(schema & "sendusername")     = "douglasproglimateste@gmail.com"
 
'Envia password
oConf.Fields(schema & "sendpassword")     = "123456"
 
oConf.Fields.Update()
 
' Envia mensagem
oMsg.Send()
 
' Retorna o status da mensagem
If Err Then
    resultMessage = "ERROR " & Err.Number & ": " & Err.Description
    Err.Clear()
Else
    resultMessage = "Mensagem enviada com sucesso !!!"
End If
 
Wscript.echo(resultMessage)