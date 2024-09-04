from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
from customScripts.readConfig import queryByNameDict
import logging
import pandas as pd 
import re
from Automation.BusinessRuleExceptions import *
import win32com.client
from pywinauto import Application
from Automation.MailboxRVS import SearchMailInbox, MoveEmailToFolder
#cd diretorio chrome
#chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\selenium\chrome-profile"

#TO-DO - Colocar Try Catch nas ações que podem ser systemError(seletores) para enviar logs com a origem do erro (nome funçao e ficheiro)

#Abrir GIO (parte ainda não desenvolvida por diferenças de ambiente)
def OpenGIO(logger:logging.Logger,dictConfig):
    #Path = queryByNameDict('PathDriverEdge',dictConfig)
    Browser_options = Options()
    Browser_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    Path= r"realvidaseguros\lib\chromedriver.exe"
    driver = webdriver.Chrome(service=Service(Path),options=Browser_options)
    LinkGIO = queryByNameDict('LinkGIO',dictConfig)
    driver.get(LinkGIO)
    time.sleep(10)
    try:
        driver.find_element_by_id("otherTileText")
        driver.maximize_window()
        logger.info("Website disponível!")
        return driver
    except Exception as e:
        logger.error(f"Website indisponível {e}")

#LoginGIO (parte ainda não desenvolvida por diferenças de ambiente)
def loginGIO(driver:webdriver.Chrome,dictConfig):
    search = driver.find_element_by_id("otherTileText")
    search.click()
    time.sleep(5)
    search = driver.find_element(By.NAME,'loginfmt')
    search.send_keys(queryByNameDict('EmailGIO',dictConfig))
    search.send_keys(Keys.RETURN)
    time.sleep(5)
    search = driver.find_element(By.NAME,'passwd')
    search.send_keys(queryByNameDict('PasswordGIO',dictConfig))
    time.sleep(5)
    search.send_keys(Keys.RETURN)
    time.sleep(15)
    SMS=input("Carregar Enter Após Login Efetuado: ")

#Navega para a página de pesquisa (tambem não utilizado ainda)
def navegarGIO(driver:webdriver.Chrome):
    print(driver.title)
    search = driver.find_element(By.XPATH,'/html/body/div[2]/nav/div/ul/li[3]/a')
    search.click()
 
#Faz a pesquisa pelo driver que lhe enviarmos (driver as in Nome, Apolice, NIF e Email))
def pesquisarGIO(driver:webdriver.Chrome,search,pesquisa:str):
    print(driver.title)
    searchButton = driver.find_element(By.XPATH,'/html/body/div[2]/div/form/div/div/div/div[6]/div/button[1]')
    driver.find_element(By.XPATH,'/html/body/div[2]/div/form/div/div/div/div[6]/div/button[2]').click()
    time.sleep(2)
    search.clear()
    search.send_keys(pesquisa)
    searchButton.click()
    time.sleep(4)
    
    searchNumEntries=driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[1]/div/label/select')
    searchNumEntries.click()
    time.sleep(1)
    searchNumPlus=driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[1]/div/label/select/option[4]')
    searchNumPlus.click()
    time.sleep(3)
    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[1]/div/div/div[1]/div/h2').click
    time.sleep(10)
 
#Extrai todos os registos da tabela de pesquisa do GIO, se houve mais que uma página ele extrai
def ScrapTableGIO(driver:webdriver.Chrome,logger:logging.Logger) -> pd.DataFrame:
    pattern = r'\d+'
    NumRegistos = driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[3]/div[1]/div').text
    NumRegistos = re.findall(pattern,NumRegistos)
    table = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[2]/div/table/tbody')
    
    headers=['Nome','TipoEntidade','NIF','Phone','Email','DOB',' ']
    table_data = []
    #Extair Info 
    #Se não houver registos a extrair, ou seja, 0 registos apresentados irá retornar uma dataframe vazia
    if max(list(map(int, NumRegistos))) > 0:
        while True:
            #'tr' equivale a uma row
            rows = table.find_elements(By.TAG_NAME,'tr')
            for row in rows:
                #'td' equivale à coluna dessa row
                cols = row.find_elements(By.TAG_NAME, 'td')
                col_data = [col.text for col in cols]
                table_data.append(col_data)
                logger.info(f'A Extrair row com os dados: {col_data}')
            #enquanto não extrair TUDO vai carregar no botão de next
            if not len(table_data) >= max(list(map(int, NumRegistos))):
                driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[3]/div[2]/div/ul/li[3]').click()
                time.sleep(2)
            else:
                break
        df=pd.DataFrame(table_data,columns=headers)
        logger.info('Extração efetuada com Sucesso!')
    else:
        df = pd.DataFrame
        logger.info('Sem Dados a Extrair!')
    return df

#Extrai o Email depois de entrar na entidade (Util apenas para a regra que tem de verificar se o email é vazio ou realvida)
def ScrapDetalhesEntidadeGIO(driver:webdriver.Chrome) -> str:
    #Nome = driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div/div/div/div/form/div/fieldset[1]/div[1]/div/input').get_attribute('value')
    #NumIF = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[2]/div[2]/div/div/div/div/form/div/fieldset[2]/div/div[4]/input').get_attribute('value')
    #headers=['Nome','NIF']
    #df = pd.DataFrame([[Nome,NumIF]],columns=headers)
    #print(df)
    #Só isto????
    Email = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[2]/div[1]/div/div/div/div[2]/div/div[1]/div[5]/input').get_attribute('value')
    return Email
#Extrai as Apolices TODAS apresentas para o cliente (ativas e inativas)
def ScrapApoliceGIO(driver:webdriver.Chrome,logger:logging.Logger) -> pd.DataFrame:
    pattern = r'\d+'
    NumRegistos = driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div/div/div/div[3]/div[1]/div').text
    NumRegistos = re.findall(pattern,NumRegistos)
    logger.info(f'A tentar extrair as {max(list(map(int, NumRegistos)))} Apólices...')
    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div/div/div/div[1]/div[1]/div/label/select').click()
    time.sleep(2)
    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div/div/div/div[1]/div[1]/div/label/select/option[4]').click()
    time.sleep(1)
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    time.sleep(10)
    #Extrair Info
    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div/div/div/div[2]/div/table/thead/tr/th[9]').click()
    time.sleep(2)
    while driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div/div/div/div[2]/div/table/thead/tr/th[9]').get_attribute('aria-sort') != 'descending':
        print(driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div/div/div/div[2]/div/table/thead/tr/th[9]').get_attribute('aria-sort'))
        driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div/div/div/div[2]/div/table/thead/tr/th[9]').click()
        time.sleep(2)

    time.sleep(10)
    table = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div/div/div/div[2]/div/table/tbody')
    rows = table.find_elements(By.TAG_NAME,'tr')
    headers=['Apolice','Versao','Produto','NomeTitular','1PessoaSegura','2PessoaSegura','Inicio','Termo','Situacao','']
    table_data = []
    #Extair Info 
    #tr - rows / td - columns logica: a cada tr extrair todos os td
    if max(list(map(int, NumRegistos))) > 0:
        while True:
            rows = table.find_elements(By.TAG_NAME,'tr')
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, 'td')
                col_data = [col.text for col in cols]
                logger.info(f'A Extrair row com os dados: {col_data}')
                table_data.append(col_data)
            if not len(table_data) == max(list(map(int, NumRegistos))):
                driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div/div/div/div[3]/div[2]/div/ul/li[5]/a').click()
                time.sleep(5)
            else:
                break
        df=pd.DataFrame(table_data,columns=headers)
    else:
        logger.info('Sem Dados para Extrair!')
        df = pd.DataFrame
    #Converter para dataframe
    #print(df)
    return df
    

def GetInfoCredorHipotecario(driver:webdriver.Chrome,logger:logging.Logger) -> str:
    CredorHipotecario =""
    parent_div = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/div[2]/div/div/div/div[1]/div[1]/div[3]')
    child_divs = parent_div.find_elements(By.XPATH, './div')
    i= 1
    for child_div in child_divs:
            try:
                label = child_div.find_element(By.TAG_NAME, 'label')
                label_text = label.text
                #print(f'Label text: {label_text}')
                if label_text == 'Credor Hipotecário:':
                    CredorHipotecario = driver.find_element(By.XPATH,f'/html/body/div[2]/div/div[3]/div[2]/div/div/div/div[1]/div[1]/div[3]/div[{i}]/label[2]').text
                    logger.info(f'Identificado {label_text} - {CredorHipotecario}')
            except:
                print('No label found in this div.') 
            i=i+1
    return CredorHipotecario

#Não terminado/implementado devido a dúvidas do ambiente produtivo
def send_email(mail, body,logger:logging.Logger,To, attachments=None):
    reply = mail.Reply()
    reply.To = 'brunofilipe.lobo@cgi.com' #Usado em dev para testesssss
    #reply.To = To
    reply.Body = body + reply.Body
    for attachment_path in attachments:
        reply.Attachments.Add(attachment_path)

    reply.Display()

    time.sleep(15)
    app = Application(backend='uia').connect(title_re='.*Message.*')
    main_window = app.window(title_re='.*Message.*')
    main_window.set_focus()
    try:
        main_window.child_window(title="Non-Business", control_type="ListItem").click_input()
    except:
        logger.info('Sem Label de Classificação')
    try:
        main_window.child_window(title="Send", control_type="Button").click_input()
        logger.info(f'Email para {To} , Enviado com Sucesso!')
    except Exception as e:
        logger.error(f'Impossibilidade em enviar o Email: {e}')
        raise Exception('Impossibilidade em enviar o Email')


def registarcontactoGIO(driver:webdriver.Chrome,logger:logging.Logger,df:pd.DataFrame,email:str):
    logger.info('A Proceder com o Registo do Contacto com o Cliente no GIO')
    #click Contactos
    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[1]/div/div/div/div[2]/div/div[2]/a[1]').click()
    time.sleep(5)
    #+#click Contactos
    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div[1]/div[2]/div/div/div[1]/button').click()
    time.sleep(5)
    #click tipificacao
    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div[2]/div[1]/div/div/div/form/div[2]/div[10]/span/span[1]/span/span[1]/span').click()
    time.sleep(5)
    #dar filtro tipificacao
    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div[2]/div[1]/div/div/div/form/div[2]/div[10]/select/option[7]').click()
    time.sleep(5)
    #click estado
    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div[2]/div[1]/div/div/div/form/div[2]/div[11]/span/span[1]/span').click()
    time.sleep(5)
    #dar filtro estado
    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div[2]/div[1]/div/div/div/form/div[2]/div[11]/select/option[3]').click()
    time.sleep(5)
    #click tipo
    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div[2]/div[1]/div/div/div/form/div[2]/div[9]/span/span[1]/span/span[1]/span').click()
    #select email
    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div[2]/div[1]/div/div/div/form/div[2]/div[9]/select/option[2]').click()
    time.sleep(5)
    desc = driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div[2]/div[1]/div/div/div/form/div[2]/div[16]/textarea')
    desc.clear()
    textodesc = f"Assunto do Email: {df.loc[0,'Subject']} \nCorpo do Email:  {df.loc[0,'Body']}\nTema NLP: {df.loc[0,'IDIntencao']}\nTemplate Resposta Enviado: {email}"
    desc.send_keys(textodesc)
    time.sleep(5)
    #click guardar
    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div[2]/div[1]/div/div/div/form/div[7]/div/button[2]').click()
    time.sleep(3)



def idAlertas(driver:webdriver.Chrome,dfInfoRegisto:pd.DataFrame,dictConfig,logger:logging.Logger):
    time.sleep(3)
    searchEmail = driver.find_element(By.XPATH,'/html/body/div[2]/div/form/div/div/div/div[4]/input')
    searchName = driver.find_element(By.XPATH,'/html/body/div[2]/div/form/div/div/div/div[1]/input')
    searchNIF = driver.find_element(By.XPATH,'/html/body/div[2]/div/form/div/div/div/div[2]/input')
    searchApolc = driver.find_element(By.XPATH,'/html/body/div[2]/div/form/div/div/div/div[5]/input')
    file_path = queryByNameDict('PathConfigIntencoes',dictConfig)
    auxfile_path= queryByNameDict('PathClassificacaoApolices',dictConfig)
    
    dfRegras = pd.read_excel(file_path,keep_default_na=False,sheet_name='IdentifEntidade')
    
    #Saber que Regras Ignorar (onde todas as columns são NA)
    dfRegrasNA = dfRegras[dfRegras.drop(columns='ID').map(lambda x: x == 'NA').all(axis=1)]
    #print(dfRegrasNA)

    if dfInfoRegisto.loc[0,'IDIntencao'] in dfRegras['ID'].values:
        rowAnalise = dfRegras.loc[dfRegras['ID'] == dfInfoRegisto.loc[0,'IDIntencao']]
    else:
        raise BusinessRuleException("ID atribuido pelo NLP não configurado!")
    logger.info(f'A utilizar a regra - {rowAnalise.values}')
    boolMatch = False

    dfRegrasAnexos = pd.read_excel(file_path,keep_default_na=False,sheet_name='Anexos') 
    if dfInfoRegisto.loc[0,'IDIntencao'] in dfRegrasAnexos[dfRegrasAnexos.drop(columns='ID').map(lambda x: x == 'Não').all(axis=1)].values:
        if not dfInfoRegisto.loc[0,'Anexos'] == 'False':
            raise BusinessRuleException('Registo contém Anexos, enquando a Regra não permite')

    for index, row in rowAnalise.iterrows():
        for col in dfRegras.columns:
            if not col == 'ID':
                if row[col] == 'NA':
                        logger.info(f'Regra para a coluna {col} não se aplica.')
                else:
                    match col:
                        case 'Email':
                            search = searchEmail
                            pesquisa = dfInfoRegisto.loc[0,'EmailRemetente']
                        case 'NIF':
                            search = searchNIF
                            pesquisa = dfInfoRegisto.loc[0,'NIF']
                        case 'Nome':
                            search = searchName
                            pesquisa = dfInfoRegisto.loc[0,'Nome']
                        case 'Apólice':
                            search = searchApolc
                            pesquisa = dfInfoRegisto.loc[0,'Apolice']
                        case _:
                            logger.info(f"Skipping Coluna {col}")
                    if pesquisa.strip().replace(' ','') == '':
                        logger.info(f'Sem Dados para Pesquisar em {col}')
                    else:
                        for p in pesquisa.split('|'):
                            logger.info(f'A Pesquisar no GIO por {col} o valor: {p.strip()}')
                            pesquisarGIO(driver,search,p.strip())
                            dfGIO = ScrapTableGIO(driver,logger)
                            #print(dfGIO['Tipo Entidade'])
                            if not dfGIO.empty:
                                if row[col] == 'Não':
                                    raise BusinessRuleException(f"A Pesquisa do Campo {col} retornou valores sendo que não é suposto. ID: {row['ID']} Regra: {row[col]}.")
                                for value in row[col].split(";"):
                                    #print(value)
                                    if any(val == value for val in dfGIO['TipoEntidade'].values):
                                        for index,rowGIO in dfGIO['TipoEntidade'].items():
                                            logger.info(f'A procurar Match entre {col} com o Tipo de Entidade extraído de {rowGIO} com a regra...')
                                            if rowGIO == value:
                                                boolMatch = True
                                                logger.info(f'Match com Regra {col}, a regra diz {row[col].replace("0","").replace(" ","")} e o registo extraído tem {rowGIO.replace("0","").replace(" ","")}')
                                                colunaMatch = col
                                                while index > 100:
                                                    index = index-100
                                                    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[3]/div[2]/div/ul/li[3]').click()
                                                    if index < 100:
                                                        break                                        
                                                driver.find_element(By.XPATH,f'/html/body/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[2]/div/table/tbody/tr[{index+1}]/td[7]/div').click()
                                                break
                                        break
                            else:
                                webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                                webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                if boolMatch:
                        break
            if boolMatch:
                break
            else:
                webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        
    boolMatchNA =False

    if not boolMatch and not any(val == dfInfoRegisto.loc[0,'IDIntencao'] for val in dfRegrasNA['ID'].values):
        logger.warn('Sem Match Com Nenhuma das Regras Definidas')
        logger.info('A Verificar se o Registo contem um Cliente RealVida')
        for index, row in rowAnalise.iterrows():
            for col in dfRegras.columns:
                if not col == 'ID':
                    match col:
                        case 'Email':
                            search = searchEmail
                            pesquisa = dfInfoRegisto.loc[0,'EmailRemetente']
                        case 'NIF':
                            search = searchNIF
                            pesquisa = dfInfoRegisto.loc[0,'NIF']
                        case 'Nome':
                            search = searchName
                            pesquisa = dfInfoRegisto.loc[0,'Nome']
                        case 'Apólice':
                            search = searchApolc
                            pesquisa = dfInfoRegisto.loc[0,'Apolice']
                    if pesquisa.strip().replace(' ','') == '':
                        logger.info(f'Sem Dados para Pesquisar em {col}')
                        boolMatchNA = True
                    else:
                        for p in pesquisa.split('|'):
                            logger.info(f'A Pesquisar no GIO por {col} o valor: {p.strip()}')
                            pesquisarGIO(driver,search,p.strip())
                            dfGIO = ScrapTableGIO(driver,logger)
                            if not dfGIO.empty:
                                raise BusinessRuleException("Registo com Cliente RealVida - Sem Match Com Nenhuma das Regras Definidas")
                            else:
                                boolMatchNA = True
                                webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                                webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()                           
    elif not boolMatch and any(val == dfInfoRegisto.loc[0,'IDIntencao'] for val in dfRegrasNA['ID'].values):
        rowAnalise = rowAnalise[rowAnalise.drop(columns='ID').map(lambda x: x == 'NA').all(axis=1)]
        dfInfoRegisto.loc[0,'IDIntencao'] = '3NA'
        
    if boolMatchNA == True:
        logger.warning('Impossibilidade de Identificação de Cliente RealVida')
        dfInfoRegisto.loc[0,'IDIntencao'] = 'NA'
        rowAnalise = dfRegras.loc[dfRegras['ID'] == dfInfoRegisto.loc[0,'IDIntencao']]

    #Identifacao Alternativa para casos de clientes sem emails registados/vazios

    dfRegrasIdentifAltern = pd.read_excel(file_path,keep_default_na=False,sheet_name='IdentifAlternativa') 
    if dfInfoRegisto.loc[0,'IDIntencao'] in dfRegrasIdentifAltern[dfRegrasIdentifAltern.drop(columns='ID').map(lambda x: x == 'Sim').all(axis=1)].values:
        dfRegrasSIM = dfRegrasIdentifAltern[dfRegrasIdentifAltern.drop(columns='ID').map(lambda x: x == 'Sim').all(axis=1)]
        row = dfRegrasSIM.loc[dfRegrasSIM['ID'] == dfInfoRegisto.loc[0,'IDIntencao']]
        for col in dfRegrasSIM.columns:
            if col == colunaMatch and row[col].values == 'Sim':
                time.sleep(5)
                EmailCliente = ScrapDetalhesEntidadeGIO(driver)
                #print(EmailCliente=='')
                #print(EmailCliente==' ')
                #print(EmailCliente=='x@x.pt')
                if not EmailCliente == '' and not EmailCliente == 'x@x.pt': #colocar config email/emails permitidos
                    raise BusinessRuleException(f'Registo contém um email: {EmailCliente}, de acordo com a Regra tem de ser @realvida ou vazio.')
    
    #Verificação de Apolices Ativas (Caso necessário)
    dfRegrasApoliceAtivas = pd.read_excel(file_path,keep_default_na=False,sheet_name='ApolAtivas')
    #Apenas entra aqui se o ID identificado tenha alguma coluna sem 'NA' else não entrará
    if not dfInfoRegisto.loc[0,'IDIntencao'] in dfRegrasApoliceAtivas[dfRegrasApoliceAtivas.drop(columns='ID').map(lambda x: x =='NA').all(axis=1)] and boolMatch:
        rowAnalise = dfRegrasApoliceAtivas[dfRegrasApoliceAtivas.drop(columns='ID').map(lambda x: (x !='NA')).any(axis=1)].loc[dfRegrasApoliceAtivas['ID'] == dfInfoRegisto.loc[0,'IDIntencao']]
        driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[1]/div/div/div/div[2]/div/div[2]/a[4]').click()
        time.sleep(5)
        dfApolicesAtivas = ScrapApoliceGIO(driver,logger)
        dfApolicesAtivas =dfApolicesAtivas[dfApolicesAtivas.map(lambda row: row == 'EM VIGOR').any(axis=1)]
        if dfApolicesAtivas.empty or dfApolicesAtivas[dfApolicesAtivas.map(lambda row: row == 'EM VIGOR').any(axis=1)].empty :
            raise BusinessRuleException('Sem Apolices Ativas')
        logger.info(f"A Proceder com Verificação às Apolices Ativas com a Regra {rowAnalise.values}")
        #print(dfApolicesAtivas)
        dfClassicaoApolices = pd.read_excel(auxfile_path,keep_default_na=False,sheet_name='Classificação Produtos')
        dfApolicesAtivas['ApoliceVersao'] = dfApolicesAtivas.apply(lambda row: row['Apolice'].split('-')[0] +'/'+row['Versao'], axis=1)
        dfClassicaoApolices['ApoliceVersao'] = dfClassicaoApolices.apply(lambda row:str(row['MODALIDADE']) + '/' + str(row['VERSAO']),axis=1)
        logger.info(f"A Aplicar Regras para as Apolices/Versões: {dfApolicesAtivas['ApoliceVersao'].values}")

        listMatches = []
    
        for i, row in rowAnalise.iterrows():
            for col in rowAnalise.drop(columns='ID').columns:
                if row[col] == 'NA':
                    logger.info(f'Não existe Regra Definida para {col}')
                else:
                    boolMatch =False
                    logger.info(f'A analisar a Regra de {col}')
                    match col:
                        case 'Modalidade/Versão Em Vigor':
                            if not row[col] == 'Todos':
                                for RegraModalidadeVersao in row[col].split('ou'):
                                    #print(RegraModalidadeVersao)
                                    if 'X' in RegraModalidadeVersao.split('/')[0]:
                                        for ApoliceVersao in dfApolicesAtivas['ApoliceVersao']:
                                            if RegraModalidadeVersao.split('/')[1] == ApoliceVersao.split('/')[1]:
                                                #print("Existe Match")
                                                boolMatch = True
                                                listMatches.append(ApoliceVersao)
                                                #break
                                    elif 'X' in RegraModalidadeVersao.split('/')[1]:
                                        for ApoliceVersao in dfApolicesAtivas['ApoliceVersao']:
                                            if RegraModalidadeVersao.split('/')[0] == ApoliceVersao.split('/')[0]:
                                                #print("existe")
                                                boolMatch = True
                                                listMatches.append(ApoliceVersao)
                                                #break
                                    elif any(RegraModalidadeVersao == ApoliceVersao for ApoliceVersao in dfApolicesAtivas['ApoliceVersao']):
                                        for ApoliceVersao in dfApolicesAtivas['ApoliceVersao']:
                                            if RegraModalidadeVersao == ApoliceVersao:
                                                #print("existe")
                                                boolMatch = True
                                                listMatches.append(ApoliceVersao)
                                                #break
                                if not boolMatch:
                                    raise BusinessRuleException('Sem Match Com Regra')
                                logger.info(f'Match com a Regra nas Apolices/Versões {listMatches}')
                            elif dfApolicesAtivas.empty:
                                raise BusinessRuleException('Sem Match Com Regra')
                            #break
                        case 'Limitação de Modalidade/Versão':
                            if not row[col] == 'Todos':
                                for RegraModalidadeVersao in row[col].split('ou'):
                                    #print(RegraModalidadeVersao)
                                    if 'X' in RegraModalidadeVersao.split('/')[0]:
                                        if any(RegraModalidadeVersao.split('/')[1] == ApoliceVersao.split('/')[1] for ApoliceVersao in dfApolicesAtivas['ApoliceVersao']):
                                            raise BusinessRuleException("")
                                    elif 'X' in RegraModalidadeVersao.split('/')[1]:
                                        if any(RegraModalidadeVersao.split('/')[0] == ApoliceVersao.split('/')[0] for ApoliceVersao in dfApolicesAtivas['ApoliceVersao']):
                                            raise BusinessRuleException("")
                                    elif any(RegraModalidadeVersao == ApoliceVersao for ApoliceVersao in dfApolicesAtivas['ApoliceVersao']):
                                            raise BusinessRuleException("")
                            elif not dfApolicesAtivas[~dfApolicesAtivas['ApoliceVersao'].isin(listMatches)].empty:
                                raise BusinessRuleException('Sem Match Com Regra')
                            boolMatch=True
                            logger.info("Nenhuma Apolice/Versão Impeditiva Detetada!")
                        case 'Produto Em Vigor':
                            if row[col] == 'Todos':
                                if dfApolicesAtivas.empty:
                                    raise BusinessRuleException("Sem Match Com Regra")
                                boolMatch=True
                                logger.info(f'Match uma vez que tem Produtos Ativos')
                            else:
                                for produto in row[col].split(';'):
                                    for ApoliceVersao in dfApolicesAtivas['ApoliceVersao']:
                                        if ApoliceVersao in dfClassicaoApolices['ApoliceVersao'].values:
                                            for i,ApoliceVersaoClass in dfClassicaoApolices.iterrows():
                                                if ApoliceVersaoClass['ApoliceVersao'] == ApoliceVersao and produto == ApoliceVersaoClass['PRODUTO']:
                                                    logger.info(f'Match com o produto da Apolice/Versao {ApoliceVersao}, existindo o produto do tipo {produto}')
                                                    boolMatch = True
                                                    break
                                if not boolMatch:
                                    raise BusinessRuleException("Sem Match com Regra")
                        case 'Produto Impeditivo Em Vigor':
                            if row[col] == 'Todos':
                                if not dfApolicesAtivas.empty:
                                    raise BusinessRuleException("Sem Match com Regra")
                                boolMatch =True
                                logger.info('Match uma vez que NÃO tem Produtos Ativos')
                            else:
                                for produto in row[col].split(';'):
                                    for ApoliceVersao in dfApolicesAtivas['ApoliceVersao']:
                                        if ApoliceVersao in dfClassicaoApolices['ApoliceVersao'].values:
                                            for i,ApoliceVersaoClass in dfClassicaoApolices.iterrows():
                                                if ApoliceVersaoClass['ApoliceVersao'] == ApoliceVersao and produto == ApoliceVersaoClass['PRODUTO']:
                                                    logger.info(f'Match com o produto da Apolice/Versao {ApoliceVersao}, existindo o produto do tipo {produto}')
                                                    raise BusinessRuleException("Sem Match com Regra")
                                logger.info("Nenhum Produto Impeditivo Detetado!")
                                boolMatch=True
                        case 'Credor Hipotecário':
                            print(f"A verificar {col} das Apolices Ativas")
                            for i, rowAA in dfApolicesAtivas.iterrows():
                                driver.find_element(By.XPATH,f'/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div/div/div/div[2]/div/table/tbody/tr[{i+1}]/td[10]/div/a').click()
                                break
                            CredorHipotecario = GetInfoCredorHipotecario(driver,logger)
                            print(CredorHipotecario)
                            if row[col] == 'Sim':
                                if CredorHipotecario == '' or CredorHipotecario == None:
                                    raise BusinessRuleException("Sem Match com a Regra")
                            if row[col] == 'Não':
                                if not CredorHipotecario == '' or not CredorHipotecario == None:
                                    raise BusinessRuleException("Sem Match com a Regra")
                            boolMatch=True
                            logger.info(f'{col} em conformidade com a Regra')
                        case _:
                            print(f'{col} ainda não configurado')

    if dfInfoRegisto.loc[0,'IDIntencao'] in dfRegrasApoliceAtivas[dfRegrasApoliceAtivas.drop(columns='ID').map(lambda x: x =='NA').all(axis=1)] and not boolMatch:
        raise BusinessRuleException('Sem Match com a Regra')
    
        
    dfEmailTemplates = pd.read_excel(file_path,keep_default_na=False,sheet_name='IDTemplates')
    rowEmails = dfEmailTemplates.loc[dfEmailTemplates['ID'] == dfInfoRegisto.loc[0,'IDIntencao']]
    email = ''
    for body in rowEmails['Template']:
        email = body
        break 
    #print(rowAnalise)
    #print(rowAnalise.drop(columns='ID').map(lambda x :x == 'NA').all(axis=1).empty)
    #print(rowAnalise.drop(columns='ID').eq('NA').all(axis=1).any())
    if not rowAnalise.drop(columns='ID').eq('NA').all(axis=1).any() and not rowAnalise.drop(columns='ID').eq('Não').all(axis=1).any():
        registarcontactoGIO(driver,logger,dfInfoRegisto,email)
        #atividade final
        driver.find_element(By.ID,'deleteData').click()

def EnviarEmail(dfInfoRegisto:pd.DataFrame,dictConfig,logger:logging.Logger):
    file_path = queryByNameDict('PathConfigIntencoes',dictConfig)
    FolderTratamentoRPA = queryByNameDict("EmailsToMove",dictConfig)
    mailbox_name =  queryByNameDict("MailboxName",dictConfig)
    dfEmailTemplates = pd.read_excel(file_path,keep_default_na=False,sheet_name='IDTemplates')
    rowEmails = dfEmailTemplates.loc[dfEmailTemplates['ID'] == dfInfoRegisto.loc[0,'IDIntencao']]
    To = dfInfoRegisto.loc[0,'EmailRemetente']
    if rowEmails.empty:
        raise BusinessRuleException(f'ID: {dfInfoRegisto.loc[0,"IDIntencao"]} Sem Template Para Responder')
    for body in rowEmails['Template']:
        mail = SearchMailInbox(logger,FolderTratamentoRPA,mailbox_name,dfInfoRegisto.loc[0,'EmailID'])
        if mail:
            send_email(mail,body,logger)
            break
        else:
            raise BusinessRuleException('Email original não encontrado para efetuar resposta')
