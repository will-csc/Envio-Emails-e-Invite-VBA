Attribute VB_Name = "Reunião"
Dim Navegador As New Selenium.ChromeDriver
Sub EMAIL_REUNIÃO()

    '--------------- DEFINIÇÃO DE VARIAVEIS ----------------
    Dim Navegador As New Selenium.ChromeDriver
    Dim Intervalo As Variant
    Dim Escolha As Long, Horario As String, DiaReunião As Date
    Dim TagSpanAzul As String, TagSpanAzulBold As String, TagSpanClose As String
    
    Set Objeto_Outlook = CreateObject("Outlook.Application")
    Set Email_Enviar = Objeto_Outlook.CreateItem(0)
    Set ws = ThisWorkbook.Sheets("REUNIÃO DE VENDAS")
    
    ProxMês = Month(Date + 30)
    
    TagSpanAmarelaBold = "<span style='background-color:Yellow; color:Red;font-size:14pt;font-family:Calibri;font-weight:bold'>"
    TagSpanAmarela = "<span style='background-color:Yellow; color:Red;font-size:14pt;font-family:Calibri;font-weight:normal'>"
    TagSpanPreta = "<span style='background-color:Yellow; color:Black;font-size:30pt;font-family:Calibri;font-weight:bold'>"
    TagSpanAzul = "<span style='background-color:White; color:navy;font-size:14pt;font-family:Calibri;font-weight:normal''>"
    TagSpanAzulBold = "<span style='background-color:White; color:navy;font-size:14pt;font-family:Calibri;font-weight:bold''>"
    TagSpanClose = "</span>"
    
    If Time() > TimeValue("12:00:00") Then
        Horario = "Boa tarde!"
    Else
        Horario = "Bom dia!"
    End If
    '------------- DIAS DA REUNIÃO --------------
    With ws.Range("O17")
        .AutoFilter Field:=2, Criteria1:=ProxMês
    End With
    
    Intervalo = ws.Range("O18:Q45")
    
    For i = 1 To UBound(Intervalo)
        If Intervalo(i, 1) = ProxMês Then
            Address = ws.Range("A12:C18").Find(Intervalo(i, 3)).Offset(0, 1).Address
            ws.Range(Address) = Intervalo(i, 2)
        End If
    Next i
    
    '------------- DESTINATARIOS --------------
    Intervalo = ws.Range("AA2:AA" & ws.Range("AA1").CurrentRegion.Rows.Count)
    For i = 1 To UBound(Intervalo)
        Destinatarios = Destinatarios & ";" & Intervalo(i, 1)
    Next i
    
    '------------- CONTAINER 1 (EMAIL) --------------
    Texto1 = TagSpanAzul & Horario & "<br><br>"
    Texto2 = TagSpanAzul & Range("A9").Value & "<br><br>"
    
    Texto3 = TagSpanAzul & Range("A10").Value & "<br><br>"
    Texto3 = Replace(Texto3, "2023", TagSpanAzul & "2023" & TagSpanClose)

    '------------- CONTAINER 3 (EMAIL) --------------
    Intervalo = ws.Range("A26:A32")
    For i = 1 To UBound(Intervalo)
        texto6 = texto6 & Intervalo(i, 1) & "<br>"
    Next i
    texto6 = Replace(texto6, "RCA’S", TagSpanAmarelaBold & "RCA’S" & TagSpanClose)
    
    '------------- CONTAINER 4 (EMAIL) --------------
    Intervalo = ws.Range("A35:A38")
    texto7 = TagSpanAzul
    For i = 1 To UBound(Intervalo)
        texto7 = texto7 & Intervalo(i, 1) & "<br>"
    Next i
    texto7 = texto7 & TagSpanClose
    texto7 = Replace(texto7, "Inicio 08:00 Termino 09:30", TagSpanPreta & "Inicio 08:00 Termino 09:30" & "<br><br>" & TagSpanClose)
    texto7 = Replace(texto7, "DEPARTAMENTOS", TagSpanAmarelaBold & "DEPARTAMENTOS" & TagSpanClose)
    texto7 = Replace(texto7, Range("E35").Value, TagSpanAmarelaBold & Range("E35").Value & TagSpanClose)
    texto7 = Replace(texto7, Range("A38").Value, TagSpanAmarelaBold & Range("A38").Value & TagSpanClose)
    
    '------------- INICIALIZAÇÃO DO EMAIL --------------
    Escolha = MsgBox("Reunião Presencial?", vbYesNo)
    If Escolha = vbNo Then
        '------------- CONTAINER 2 (EMAIL) --------------
        Texto4 = TagSpanAmarelaBold & "<b>" & "Segue as datas:" & "<b>" & "<br><br>" & TagSpanClose
        texto5 = TagSpanAzulBold & Range("A20").Value & "<br>" & Range("A21").Value & "<br>" & Range("A22").Value _
        & "<br>" & Range("A23").Value & "<br><br>" & TagSpanClose
        texto6 = TagSpanAzul & texto6 & "<br><br>" & TagSpanClose
        
        With Email_Enviar.Display
                Email_Enviar.To = Destinatarios
                Email_Enviar.Subject = Range("A5").Value
                Email_Enviar.htmlbody = Texto1 & Texto2 & Texto3 & Texto4 & texto5 _
                & texto6 & texto7 & Email_Enviar.htmlbody
        End With
        DiaReunião = ws.Range("D14")
    Else
        DiaReunião = InputBox("Digite o dia da reunião, Exemplo = 11/08/2023")
        texto7 = Replace(texto7, "Inicio 08:00 Termino 09:30", "Inicio 08:00 Termino 12:30")
        Texto4 = TagSpanAzul & "Informamos que a Reunião de Vendas de " & TagSpanClose & TagSpanAmarela _
        & UCase(MonthName(Month(Date + 30))) & "/" & Year(Date + 30) & " SERÁ PRESENCIAL." & TagSpanClose & "<br><br>"
        texto5 = TagSpanAzul & "E sera realizada em um único dia, sendo ele " & TagSpanAmarela & DiaReunião & TagSpanClose _
        & " com todas as equipes presentes C1,C2 e C5." & TagSpanClose & "<br><br>"
        With Email_Enviar.Display
                Email_Enviar.To = Destinatarios
                Email_Enviar.Subject = Range("A5").Value
                Email_Enviar.htmlbody = Texto1 & Texto2 & Texto3 & Texto4 & texto5 & texto7 & Email_Enviar.htmlbody
        End With
    End If
    
    '------------- CRIAÇÃO DE PASTAS --------------
    Dim Fso As Object
    Dim CaminhoGestores As String, CaminhoVendas As String
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    CaminhoGestores = "C:\Users\wcscarvalho\Dropbox\Trip\TIME GESTORES\REUNIÃO DE VENDAS SKYPE\" & Year(Now()) _
    & "\" & Month(Now()) + 1 & ". " & UCase(MonthName(Month(Now()) + 1))
    CaminhoVendas = "C:\Users\wcscarvalho\Dropbox\1.V1 STEFANI X VENDAS\1  AA ROTINA HOME-OFF\Vendas\ROTINA REUNIAO VENDAS\ADM REUNIOES DE VENDAS\" _
    & Year(Now()) & "\" & Month(Now()) + 1 & ". " & UCase(MonthName(Month(Now()) + 1))
    
    If Not Fso.FolderExists(CaminhoGestores) Then
        MkDir CaminhoGestores
        For Each Dpto In Array("CD", "RECEBIMENTO", "CQ", "LOGISTICA", "FINANCEIRO")
            MkDir CaminhoGestores & "\" & Dpto
        Next Dpto
    End If
    If Not Fso.FolderExists(CaminhoVendas) Then
        MkDir CaminhoVendas
    End If
    
    Set Objeto_Outlook = Nothing
    Set Email_Enviar = Nothing
    
    '------------- CHAMAR E-MAIL E INVITE --------------
    Call EmailRH(Escolha, Horario, TagSpanAzul, TagSpanAzulBold, TagSpanClose, DiaReunião)
    Call InvitesReunião(Escolha, DiaReunião)
    
End Sub
Sub EmailRH(Escolha As Long, Horario As String, TagSpanAzul As String, TagSpanAzulBold As String, TagSpanClose As String _
            , DiaReunião As Date)
    
    '---------- DEFINIÇÃO VARIAVEIS P/ EMAIL ------------
    Dim Intervalo As Variant
    
    Set Objeto_Outlook = CreateObject("Outlook.Application")
    Set Email = Objeto_Outlook.CreateItem(0)
    
    '------------- PESSOAS --------------
    Intervalo = Range("AA2:AA" & Range("AA1").CurrentRegion.Rows.Count)
    For i = 1 To UBound(Intervalo)
        If InStr(1, Intervalo(i, 1), "bufalo") = 0 Then
            Pessoas = Pessoas + 1
        End If
    Next i
    Pessoas = Pessoas + 6
    
    '------------- CRIAÇÃO DO EMAIL ---------------
    If Escolha = vbYes Then
        Texto1 = TagSpanAzul & Horario & "<br><br><b>@DP</b>, Conforme todas reuniões presenciais realizadas todo trimestre" & _
        ", peço para que os senhores reservem a sala de treinamento das 06h30 até 12h30 do dia " _
        & DiaReunião & "<br>Se possível solicitar Café e Leite para " & Pessoas & " Pessoas.<br><br>" & TagSpanClose
        Texto2 = TagSpanAzul & "Disponibilizar o Café as 07:00 em frente ao refeitorio, Obrigado!." & TagSpanClose
        With Email.Display
            Email.To = "dp@produtosbufalo.com.br;dp.aux@produtosbufalo.com.br;bufalo@llinea.com.br"
            Email.Cc = "vendas1@produtosbufalo.com.br"
            Email.Subject = UCase(Range("A5").Value) & " | RESERVA SALA E CAFÉ"
            Email.htmlbody = Texto1 & Texto2 & Email.htmlbody
        End With
    Else
        Texto1 = TagSpanAzul & Horario & " DP, Conforme todas reuniões realizadas todos os meses" & _
        ", peço para que os senhores reservem a sala de reunião das 08h00 até 09h30 conforme dias e equipes abaixo<br><br>"
        Texto2 = TagSpanAzulBold & Range("A20").Value & "<br>" & Range("A21").Value & "<br>" & Range("A22").Value _
        & "<br>" & Range("A23").Value & "<br><br>" & TagSpanClose
        Texto3 = TagSpanAzul & "Aguardo o retorno e aprovação, Obrigado!." & TagSpanClose
        With Email.Display
            Email.To = "dp@produtosbufalo.com.br;dp.aux@produtosbufalo.com.br"
            Email.Cc = "vendas1@produtosbufalo.com.br"
            Email.Subject = UCase(Range("A5").Value) & " | RESERVA SALA ADM"
            Email.htmlbody = Texto1 & Texto2 & Texto3 & Email.htmlbody
        End With
    End If
    
End Sub

Sub InvitesReunião(Escolha As Long, DiaReunião As Date)

    Navegador.Start
    Navegador.Get "https://calendar.google.com/calendar/u/0/r/day/2023/10/9"
    
    '------------ VARIAVEIS ------------
    Dim Dias As Variant
    
    '---------- LOGANDO SITE -----------------
    On Error GoTo Erro:
    Set FazerLogin = Navegador.FindElementsByCss(".button.button--low").Item(3)
    FazerLogin.Click
    
    Set Email = Navegador.FindElementById("identifierId")
    Email.SendKeys "vendas@produtosbufalo.com.br"
    Navegador.SendKeys Navegador.Keys.Enter
    
    Navegador.Wait 3500
    Set Senha = Navegador.FindElementByCss(".whsOnd.zHQkBf")
    Senha.SendKeys "bufven21"
    Navegador.SendKeys Navegador.Keys.Enter
    
    '---------- LOCALIZANDO DIAS DO INVITE -----------------
    Mês = Application.WorksheetFunction.Proper(MonthName(Range("O15").Value))
    If Escolha <> vbYes Then
        Dados = Range("A14:B17").Value
    Else
        Dados = Range("A14:B14").Value
        Dados(1, 1) = "PRESENCIAL " & UCase(MonthName(Month(Date + 30))) & "/" & Year(Date + 30)
    End If
    AssuntoReunião = "REUNIÃO DE VENDAS - "
    
    '------------ LOOPING ---------------
    Navegador.Wait 2500
    Set MêsGoogle = Navegador.FindElementByXPath("//span[@class='r4nke ']")
    If InStr(1, MêsGoogle.Text, LCase(Mês)) = 0 Then
        Set ProxMês = Navegador.FindElementByXPath("//button[@aria-label='Próximo mês']")
        ProxMês.Click
        Set DiasNav = Navegador.FindElementsByXPath("//div[@class='r4nke ']")
        For i = 1 To UBound(Dados)
            For Each Dia In DiasNav
                If Dia.Text = Trim(Dados(i, 2)) Then
                    Navegador.Wait 7500
                    '----------------- DEFINIÇÃO DO DIA -----------------
                    Dia.Click
                    MsgBox ("Aperte no site")
                    Navegador.Wait 9500
                    SendKeys "{ENTER}", True
                    
                    '--------------- ASSUNTO INVITE -----------------
                    Set Assunto = Navegador.FindElementByCss(".Ufn6O.shdZ7e").FindElementByTag("Input")
                    Assunto.Click
                    Assunto.SendKeys AssuntoReunião & Dados(i, 1)
                    
                    '------------ DEFINIÇÃO HORARIOS ----------------
                    Set DefHor = Navegador.FindElementByCss("[jslog='49541; track:JIbuQc']")
                    DefHor.Click
                    
                    Set Horarios = Navegador.FindElementByXPath("//div[@class='i04qJ']")
                    Set Horario1 = Horarios.FindElementByCss("[data-key='startTime']")
                    Navegador.Wait 5000
                    Navegador.SendKeys "08:00"
                    Navegador.SendKeys Navegador.Keys.Enter
                    
                    Navegador.Wait 2000
                    Set Horarios = Navegador.FindElementByXPath("//div[@class='Z5RD1e XAsDAf']")
                    Set Horario2 = Horarios.FindElementByXPath("//input[@aria-label='Horário de término']")
                    Horario2.Click
                    Navegador.Wait 500
                    If Escolha = vbYes Then
                        Navegador.SendKeys "12:30"
                    Else
                        Navegador.SendKeys "09:30"
                    End If
                    Navegador.SendKeys Navegador.Keys.Enter
                    
                    '------------ CONVIDADOS -----------------
                    Set Convidados = Navegador.FindElementByXPath("//div[@jsname='Ik8OMb']")
                    Convidados.Click
                    Navegador.Wait 1000
                    For Each Dpto In Array("CD@PRO", "RECEBIMENTO@PRO", "FINANCEIRO@PRO", "FINANCAS@PRO", "SGQ@", "VENDAS1@PROD")
                        Navegador.SendKeys Dpto
                        Navegador.Wait 1000
                        Navegador.SendKeys Navegador.Keys.Enter
                    Next
                    Equipe = Right(Dados(i, 1), 2)
                    Limite = Range("AB1000").End(xlUp).Row
                    
                    Dim A As Long
                    If Escolha = vbYes Then
                        For A = 2 To Limite
                            Mail = Cells(A, 28).Value
                            If Mail <> "" Then
                                Navegador.SendKeys Mail
                                Navegador.Wait 1500
                                Navegador.SendKeys Navegador.Keys.Enter
                            End If
                        Next A
                    End If
                    
                    Edge = ThisWorkbook.Sheets("RELATORIOS DIARIOS").Range("a1").End(xlDown).Row
                    If Escolha = vbNo Then
                        If Equipe = "C5" Then
                            Gestor = "comercial5@produtosbufalo.com.br"
                            Navegador.SendKeys Gestor
                            Navegador.Wait 1000
                            Navegador.SendKeys Navegador.Keys.Enter
                            
                            EmailsEquipes = Get_EmailsEquipe(Equipe) 'EMAILS DAS EQUIPES
                            For Each Email In EmailsEquipes
                                Navegador.SendKeys Email
                                Navegador.Wait 1000
                                Navegador.SendKeys Navegador.Keys.Enter
                            Next Email
                        ElseIf Equipe = "C1" Or Equipe = "C2" Then
                            Gestor = "comercial6@produtosbufalo.com.br"
                            Navegador.SendKeys Gestor
                            Navegador.Wait 1000
                            Navegador.SendKeys Navegador.Keys.Enter
                            
                            EmailsEquipes = Get_EmailsEquipe(Equipe) 'EMAILS DAS EQUIPES
                            For Each Email In EmailsEquipes
                                Navegador.SendKeys Email
                                Navegador.Wait 1000
                                Navegador.SendKeys Navegador.Keys.Enter
                            Next Email
                        Else
                            For Each Comercial In Array("COMERCIAL3@PRO", "COMERCIAL4@PRO", "COMERCIAL5@PRO", "COMERCIAL6@PRO", _
                            "COMERCIAL9@PRO")
                                Navegador.SendKeys Comercial
                                Navegador.Wait 1500
                                Navegador.SendKeys Navegador.Keys.Enter
                            Next
                        End If
                    Else
                        For Each Comercial In Array("COMERCIAL5@PRO", "COMERCIAL6@PRO")
                            Navegador.SendKeys Comercial
                            Navegador.Wait 1500
                            Navegador.SendKeys Navegador.Keys.Enter
                        Next
                    End If
                    
                    '------------ ENVIA OK -------------
                    Set Button = Navegador.FindElementByCss(".VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-k8QpJ.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.nCP5yc.AjY5Oe.DuMIQc.LQeN7.pEVtpe")
                    Button.Click
                    Navegador.Wait 1500
                    Set Enviar = Navegador.FindElementByCss(".uArJ5e.UQuaGc.kCyAyd.l3F1ye.ARrCac.HvOprf.fY7wqd.M9Bg4d")
                    Enviar.Click
                    
                    If Escolha = vbYes Then
                        Set Enviar = Navegador.FindElementByCss(".uArJ5e.UQuaGc.kCyAyd.l3F1ye.ARrCac.HvOprf.evJWRb.M9Bg4d")
                        Enviar.Click
                    End If
                    
                    Exit For
                End If
            Next Dia
            If Escolha = vbYes Then Exit For
        Next i
        
    End If
    
    Exit Sub
    
Erro:
    Navegador.Wait 1000
    Limite = Limite + 1
    If Limite > 10 Then Stop
    Resume
End Sub

Function Get_EmailsEquipe(Equipe As String) As Collection
    
    '-------------- DEFINIÇÃO VARIAVEIS --------------
    Set ws_emails = ThisWorkbook.Sheets("RELATORIOS DIARIOS")
    EmailsGerais = ws_emails.Range("A1").CurrentRegion.Value
    Set Get_EmailsEquipe = New Collection
    
    '-------------- LOOPING ---------------
    For i = 2 To UBound(EmailsGerais)
        If EmailsGerais(i, 1) = Equipe And EmailsGerais(i, 4) <> "" Then
            Get_EmailsEquipe.Add EmailsGerais(i, 4)
        End If
    Next i
    
End Function

