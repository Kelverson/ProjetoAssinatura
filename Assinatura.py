import requests
import json
import win32com.client
import os
import time
import win32com.client as win32
import gc
import winreg as reg

conteudo = []
repost = False

try:
    #busco txt 
    url = fr'http://186.225.26.249:8100/home/file.txt'
    response = requests.get(url)
    response.raise_for_status()  # Levanta um HTTPError se o status não for 200

except requests.exceptions.HTTPError as e:
    repost = True
    print(f'erro ao tentar acesser o recurso: {e}')

except requests.exceptions.RequestException as e :
    repost = True
    print(f'erro ao tentar acesser o recurso: {e}')

if(repost != True):

    conteudo = response.text     # Obtém o conteúdo do arquivo como string
    linhas = conteudo.splitlines()  
    # Alternativa usando a variável de ambiente
    username_env = os.environ.get('USERNAME')
    # Obter o caminho do perfil do usuário
    userprofile = os.environ.get('USERPROFILE')
    # Obter o nome do usuário da máquina
    username = userprofile
    #print(f'UserName: {username_env}')
    #print(f'UserProfile: {userprofile}')
    # Inicializa a aplicação Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    # Obtém a sessão MAPI
    session = outlook.Session
    # Obtém todas as contas do Outlook
    accounts = session.Accounts
    emails = []
    
    #if(email == 'newline.ind.br'):
    # Itera sobre cada conta e imprime informações
    for account in accounts:
        
        #print('teste',account.DisplayName)
        email = account.DisplayName.split('@')[1]
        #print('teste',email)
        if(email == 'newline.ind.br'):

            emails.append(account.DisplayName)
            print(f"Nome da Conta: {account.DisplayName}")
            #print(f"SMTP Address: {account.SmtpAddress}")
            print("-" * 40)
            client_id = ""
            tenant_id = ""
            #6meses/validade
            client_secret=  linhas[0]
            msal_authority = f"https://login.microsoftonline.com/{tenant_id}"
            msal_scope = ["https://graph.microsoft.com/.default"]
            from  msal import ConfidentialClientApplication
            msal_app = ConfidentialClientApplication(
                client_id=client_id,
                client_credential=client_secret,
                authority=msal_authority,
            )
            result = msal_app.acquire_token_silent(
                scopes=msal_scope,
                account=None,
            )
            if not result:
                result = msal_app.acquire_token_for_client(scopes=msal_scope)
            if "access_token" in result:
                access_token = result["access_token"]
            else: 
                raise Exception("No Acess token found")
            #print(access_token)
            headers= {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }
            for email in emails :
                #email = "kelverson.silva@newline.ind.br"
                #print("Teste: ",email)
                response = requests.get(
                    url=f"https://graph.microsoft.com/beta/users('{email}')/profile",
                    #url=f"https://graph.microsoft.com/beta/users('kelverson.silva@studioluce.com.br')/profile",
                    headers=headers,    
                )
                # Convertendo a resposta JSON para um dicionário Python
                response_dict = response.json()
                #print(response_dict)
                data = response.json()
                #print(data)
                # Inicializar dicionário filtrado
                filtered_data = {
                    "postalCode": None,
                    "state": None,
                    "city": None,
                    "street": None,
                    "department": None,
                    "displayName": None,
                    "number":None,
                    "ramal":None,
                }
                # Convertendo o dicionário Python para uma string JSON com codificação UTF-8
                #response_json = json.dumps(data, indent=4, ensure_ascii=False)
                # Imprimindo a string JSON
                #print(response_json)
                # Extrair dados de 'positions'
                if 'positions' in data:
                    for position in data['positions']:
                        if 'detail' in position and 'company' in position['detail'] and 'address' in position['detail']['company']:
                            address = position['detail']['company']['address']
                            filtered_data["postalCode"] = address.get("postalCode")
                            filtered_data["state"] = address.get("state")
                            filtered_data["city"] = address.get("city")
                            filtered_data["street"] = address.get("street")
                            filtered_data["department"] = position['detail']['company'].get("department")
                            filtered_data["officeLocation"] = position['detail']['company'].get("officeLocation")
                            #print(filtered_data["officeLocation"])
                # Extrair 'displayName' de 'names'
                if 'names' in data:
                    for name in data['names']:
                        if 'displayName' in name:
                            filtered_data["displayName"] = name.get("displayName")
                            break  # Assumindo que precisamos apenas do primeiro nome
                if 'phones' in data:
                    for phone in data['phones']:
                        #print(phone)
                        if phone['type'] == 'business' and 'number' in phone:
                            filtered_data['number'] = phone.get("number")
                        elif phone['type'] == 'other' and 'number' in phone:
                            filtered_data['ramal'] = phone.get("number")
                # Convertendo o dicionário Python para uma string JSON com codificação UTF-8
                response_json = json.dumps(filtered_data, indent=4, ensure_ascii=False)
                filtered_data = json.loads(response_json)
                # Imprimindo a string JSON
                #print(filtered_data)
                # Inicializar o Word com várias tentativas
                def initialize_word():
                    attempts = 3
                    for attempt in range(attempts):
                        try:
                            word = win32.gencache.EnsureDispatch('Word.Application')
                            return word
                        except Exception as e:
                            print(f"Tentativa {attempt + 1} falhou: {e}")
                            time.sleep(2)
                    raise Exception("Não foi possível inicializar o Word após várias tentativas.")
                # Função para liberar objetos COM
                def release_com_object(obj):
                    if obj:
                        del obj
                        gc.collect()
                # Inicializar o Word
                word = initialize_word()
                word.Visible = False
                try:
                    # Criar um novo documento
                    doc = word.Documents.Add()
                    # Adicionar um delay para garantir que o documento foi criado
                    time.sleep(1)
                    # Inserir uma imagem no início do documento
                    #image_path = os.path.join(f'https://faq.newline.ind.br/hubfs/Assinatura%20Email/Logo.png')
                    image_path = os.path.join(f'{linhas[1]}')
                    range_ = doc.Range(0, 0)
                    inline_shape = doc.InlineShapes.AddPicture(image_path, False, True, range_)
                    # Definir o tamanho da imagem
                    inline_shape.Width = 425  # largura em pontos
                    inline_shape.Height = 80  # altura em pontos
                    # Adiciona um hyperlink à imagem
                    hyperlink_url = "http://www.newline.ind.br"
                    doc.Hyperlinks.Add(Anchor=inline_shape, Address=hyperlink_url)
                    # Inserir uma quebra de parágrafo após a imagem
                    range_ = doc.Range(inline_shape.Range.End, inline_shape.Range.End)
                    range_.InsertParagraphAfter()
                    # Adicionar um título abaixo da imagem
                    range_ = doc.Range(doc.Content.End - 1, doc.Content.End)
                    range_.Text = filtered_data["displayName"]
                    range_.Font.Name = 'Verdana'
                    range_.Font.Size = 15
                    #range_.Font.Bold = True
                    range_.Font.Color = 6579816 
                    range_.ParagraphFormat.LineSpacing = 11
                    range_.ParagraphFormat.SpaceAfter = 0
                    #range_.SpaceBefore = 0
                    range_.InsertParagraphAfter()
                    # Adicionar um parágrafo abaixo do título
                    range_ = doc.Range(doc.Content.End - 1, doc.Content.End)
                    range_.Text = f'{filtered_data["department"]}\n'
                    #range_.Font.family = 'Verdana'
                    range_.Font.Size = 9.5
                    range_.Font.Bold = False
                    range_.InsertParagraphAfter()
                    # Adicionar um parágrafo abaixo do título
                    range_ = doc.Range(doc.Content.End - 1, doc.Content.End)
                    if(filtered_data["ramal"] != None):
                        range_.Text = f'Telefone: +55 {filtered_data["number"]} | Ramal: {filtered_data["ramal"]}\n'
                    else:
                        range_.Text = f'Telefone: +55 {filtered_data["number"]}\n'
                    range_.Font.family = 'verdana'
                    range_.Font.Size = 9.5
                    range_.Font.Bold = False
                    range_.InsertParagraphAfter()
                    # Adicionar um parágrafo abaixo do título
                    range_ = doc.Range(doc.Content.End - 1, doc.Content.End)
                    if(filtered_data["officeLocation"] != ''):
                        range_.Text = filtered_data["officeLocation"]
                    else:
                        range_.Text = "FÁBRICA SUZANO"
                    #range_.Font.family = 'verdana'
                    range_.Font.Size = 9.5
                    range_.Font.Bold = False
                    range_.InsertParagraphAfter()
                    # Adicionar um parágrafo abaixo do título
                    range_ = doc.Range(doc.Content.End - 1, doc.Content.End)
                    range_.Text = filtered_data["street"]
                    #range_.Font.family = 'verdana'
                    range_.Font.Size = 9.5
                    range_.Font.Bold = False
                    range_.InsertParagraphAfter()
                    # Adicionar um parágrafo abaixo do título
                    range_ = doc.Range(doc.Content.End - 1, doc.Content.End)
                    range_.Text = filtered_data["city"]+"-"+filtered_data["state"]+" | "+filtered_data["postalCode"]+" | Brasil"
                    #range_.Font.family = 'verdana'
                    range_.Font.Size = 9.5
                    range_.Font.Bold = False
                    range_.InsertParagraphAfter()
                     # Adicionar texto misto com hyperlink abaixo do parágrafo
                    range_hyperlink_text = doc.Range(doc.Content.End - 1, doc.Content.End)
                    #range_.Font.family = 'verdana'
                    range_.Font.Size = 9.5
                    range_hyperlink_text.Text = "Sites: "
                    range_hyperlink_text.InsertAfter(" ")
                    # Adicionar o hyperlink
                    range_hyperlink = doc.Range(doc.Content.End - 1, doc.Content.End)
                    range_.Font.Size = 9.5
                    #range_.Font.family = 'verdana'
                    hyperlink_text = "www.newline.ind.br"
                    hyperlink_url = "http://www.newline.ind.br"
                    doc.Hyperlinks.Add(Anchor=range_hyperlink, Address=hyperlink_url, TextToDisplay=hyperlink_text)
                    range_hyperlink.InsertParagraphAfter()
                    range_hyperlink.Font.Color = win32.constants.wdColorBlue  # Definindo a cor como azul
                    # Definir o caminho onde salvar o documento
                    diretorio = os.path.join(rf'{username}\AppData\Roaming\Microsoft\Signatures')
                    #print(f'Diretorio',{diretorio})
                    caminho = os.path.join(diretorio, f'Assinatura ({email}).htm')
                    # Escapar caracteres especiais
                    caminho = rf'"{caminho}"'
                    caminho = str(caminho)
                    # Verificar se o diretório existe, se não, criar
                    if not os.path.exists(diretorio):
                        os.makedirs(diretorio)
                    # Usar o método SaveAs2 em vez de SaveAs
                    doc.SaveAs2(caminho, FileFormat=win32.constants.wdFormatHTML)

                finally:
                    # Fechar o documento e o Word
                    doc.Close(False)
                    word.Quit()
                    # Liberar os objetos COM
                    release_com_object(range_)
                    release_com_object(range_hyperlink_text)
                    release_com_object(range_hyperlink)
                    release_com_object(doc)
                    release_com_object(word)