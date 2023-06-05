from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import re 
from openpyxl import Workbook
import time

#string de comparação de email
regexEmail = re.compile(r'(qualquerString)?@(qualquer)?')
regexSite = re.compile(r'(http(s)?://)?(www)?.(qualquerstring)?.(com)(.br)?')

#função q retorna a conexão da pagina web
def drive(url):
    options = Options()
    #options.headless = True
    options.add_argument("--window-size=3840,2160")

    DRIVER_PATH = '/path/to/chromedriver'
    driver = webdriver.Chrome(options=options,executable_path=DRIVER_PATH)

    driver.get(url)
    driver.maximize_window()
    
    return driver


#pagina web principal
driver = drive("https://app.informamarkets.com.br/event/hospitalar/exhibitors/RXZlbnRWaWV3XzM5ODM5Mw==")

#laço de repetição para gerar todos os links do site
x = 0
while x <= 100:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(5)
    x+=1

#capturando a parte onde estão todos os links
pag_principal = driver.find_element(By.XPATH, '/html/body/div[2]/div[3]/div/div/div[2]/div/div/div')

#gerando um array com todos os links
links_elements = pag_principal.find_elements(By.TAG_NAME, "a")
links = []
for links_elements_unidades in links_elements:
    links.append(links_elements_unidades.get_attribute('href'))
time.sleep(2)

#setando variaveis que serão usadas no laço de repetição
num = 2
wb = Workbook()
ws1 = wb.worksheets[0]

ws1['A1'] = "Company"
ws1['B1'] = "Country"
ws1['C1'] = "First_name"
ws1['D1'] = "Last_name"
ws1['E1'] = "Name"
ws1['F1'] = "Email"
ws1['G1'] = "Job_Title"
ws1['H1'] = "Phone"
ws1['I1'] = "Industry/Main Opertation"
ws1['J1'] = "Site"

#laço de repetição para processar todos os links
for string in links:
    
    #capturando todas as paginas dos links 
    driver = drive(string)
    
    #extraindo informação da pagina do site
    #capturando company
    try:
        company = driver.find_element(By.XPATH, '//div/h1[@class="style__Name-cmp__sc-1wv3da6-7 YbzUq"]').text
    except:
        company = ""

    #descendo a pagina para gerar o resto das informações
    driver.execute_script('window.scroll(0, 200)')

    #capturando país
    try:
        pais = driver.find_element(By.XPATH, '//div[@class="chip__InlineList-ui__sc-st9ik3-1 favgFZ"]/span').text
    except:
        pais = ""

    #capturando todos os detalhes de contatos
    detalhesContato = []
    try:
        detalhesContato = driver.find_elements(By.XPATH, '//div[@class="item-with-icon__Wrapper-ui__sc-12jykqu-1 gOuAaF"]/div[@class="item-with-icon__Item-ui__sc-12jykqu-2 iBOaUN"]/a[@class="button__Wrapper-ui__sc-a2a0dz-0 inUubk"]/span[@class="button__Content-ui__sc-a2a0dz-3 jssLeg"]/span')
    except:
        detalhesContato = []
        
    #extraindo e classificando todas as informações de contato
    telefone = []
    email = []
    site = []
    endereco = []
    for info in detalhesContato:
        try:
            import phonenumbers
            telefone_ajustado = phonenumbers.parse(info.text)
            telefone.append(info.text)
        except:
            if(re.search(regexEmail,info.text)):
                email.append(info.text)
                
            elif(re.search(regexSite,info.text)):
                site.append(info.text)
                
            else:
                endereco.append(info.text)

    #descendo até o final da pagina
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    
    linksPerfil = []
    
    #coletando todos os links dos perfis da equipe
    linksPerfilBruto = driver.find_elements(By.XPATH, '//div[@class="members__CardWrapper-ea__sc-17s92m9-0 gkcigA"]/a')
    
    for linkPerfilBruto in linksPerfilBruto:
        linksPerfil.append(linkPerfilBruto.get_attribute('href'))
    
    
    #coletando os dados do membro da equipe
    for linkPerfil in linksPerfil:
        
        #acessando site do perfil
        driverperfil = drive(linkPerfil)
        
        time.sleep(5)
        
        #capturando nome do perfil
        try:
            nome = driverperfil.find_element(By.XPATH, '//div[@class="style__HeadWrapper-cmp__sc-1s7e137-0 dvdDLM"]/h2[@class="style__Name-cmp__sc-1s7e137-1 jhjTCw"]').text
        except:
            nome = ""
            
        job_title = ""
        industry = ""
        
        #extraindo e classificando todas as informações do perfil
        try:
            todosOsDados = driverperfil.find_elements(By.XPATH, '//div[@class="style__Wrapper-cmp__sc-37a2ry-0 fOjanj"]')

            for dadosElements in todosOsDados:
                
                dadosJuntos = dadosElements.text.split('\n')
                print(dadosJuntos)
                if dadosJuntos[0] == "Ramo de atividade":
                    if( type(dadosJuntos[1]) == str):
                        industry = dadosJuntos[1]
                        print("industry: " + industry)
                    else:
                        for dadosJuntosMaisde1 in dadosJuntos[1]:
                            if industry != "":
                                industry += "\n"
                            industry += dadosJuntosMaisde1
                            print("industry 2: " + industry)
                elif dadosJuntos[0] == "Cargo":
                    if(type(dadosJuntos[1]) == str):
                        job_title = dadosJuntos[1]
                        print("Job Title" + job_title)
                    else:
                        for dadosJuntosMaisde1 in dadosJuntos[1]:
                            if job_title != "":
                                job_title += "\n"
                            job_title += dadosJuntosMaisde1
                            print("Job Title2: " + job_title)
        except:
            print("")
        
        #fechando conexão
        driverperfil.close()
    
        #setando as informações na planilha
        ws1['A'+str(num)] = company
        print(company)
        
        ws1['B'+str(num)] = pais
        print(pais)
        
        todosNome = nome.split(' ')
        ws1['C'+str(num)] = todosNome[0]
        sobrenome = ""
        for splitNomes in todosNome[1:]:
            if sobrenome != "":
                sobrenome += " "
            sobrenome += splitNomes
        
        ws1['D'+str(num)] = sobrenome
        ws1['E'+str(num)] = nome
        print(nome)
        
        todosemail =''
        for emailUnidade in email:
            if todosemail != "":
                todosemail += ","
            todosemail += emailUnidade
        
        ws1['F'+str(num)] = todosemail
        print(todosemail)
        
        ws1['G'+str(num)] = job_title
        print(job_title)
        
        todostelefone = ''
        for telefoneUnidade in telefone:
            if todostelefone != "":
                todostelefone += ","
            todostelefone += telefoneUnidade
        
        ws1['H'+str(num)] = todostelefone
        print(todostelefone)
        
        ws1['I'+str(num)] = industry
        print(industry)
        
        todosSites =''
        for siteUnidade in site:
            if todosSites != "":
                todosSites += ","
            todosSites += siteUnidade
        
        print(todosSites)
        ws1['J'+str(num)] = todosSites

        num += 1
        
    #fechando conexão
    driver.close()

    #salvando a planilha gerada
    wb.save(filename = './exemplo.xlsx')