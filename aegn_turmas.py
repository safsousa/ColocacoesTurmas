
import dados_excel
import processamento_excel
import ordenacao
import variaveis
from datetime import datetime

ANO_ATUAL = 2022
N_ALUNOS_TURMA_REDUZIDA = 20
N_ALUNOS_TURMA = 24

op = variaveis.get_op()
lista_escolas = variaveis.get_escolas()
lista_colocados = variaveis.get_listaColocados()
lista_ipss = variaveis.get_ipss()

#Imprime a lista de alunos colocados
def imprime_colocados(lista_colocados):
    for key,value in lista_colocados.items():
        print(str(lista_escolas[key][0])+ " - " +key )
        print(value)

#Verifica se o aluno é condicional
def is_condicional(data_nascimento):
        
    date_time_obj = datetime.strptime(data_nascimento, '%d-%m-%Y')
    
    ano = date_time_obj.year
    mes = date_time_obj.month
    dia = date_time_obj.day
    
    idade = ANO_ATUAL - ano
    
    if idade == 6 and mes==9 and dia<15:
        return 0
    elif idade == 6 and mes==9 and dia>15:
        return 1
    elif idade == 6 and mes>9:
        return 1
    else:
        return 0


#Coloca os alunos na primeira opção
def coloca_1opcao(lista_alunos):
    
    lista_colocados = variaveis.get_listaColocados()
    
    for key,value in lista_alunos.items():
          
        if lista_colocados.get(value[7]) != None:
            #o aluno quer uma escola do agrupamento
            if is_condicional(value[1]):
                #o aluno é condicional
                y = list(lista_colocados["alunos_condicionais"])
                matricula = (key,)
                dados_aluno = matricula + (lista_alunos.get(key))
                y.append( dados_aluno)
                lista_colocados["alunos_condicionais"] = tuple(y)
            else:
                #o aluno vai ser colocado na escola que pretende
                y = list(lista_colocados[value[7]])
                matricula = (key,)
                dados_aluno = matricula + (lista_alunos.get(key))
                y.append( dados_aluno)
                lista_colocados[value[7]] = tuple(y)
        else:
            # o aluno quer sair do agrupamento
            y = list(lista_colocados["fora"])
            y.append(key)
            lista_colocados["fora"] = tuple(y)   
    
    return lista_colocados
    

#Retorna o número de alunos nees
def qt_nees(lista_alunos):
    n_alunos = 0
    for x in lista_alunos:
        if ordenacao.is_reduzturma(x):
            n_alunos = n_alunos + 1
    return n_alunos

#Altera o número máximo de alunos na escola, tendo em conta os nees
def altera_alunos_por_escola(lista_colocados):
    #altera o número de alunos por turma
    print("Números de alunos por escola devido aos nees")
    for key,value in lista_colocados.items():
        
        if key != 'fora' and key != 'alunos_condicionais':        
            #vai buscar o número de turmas da escola
            #vai buscar o número de turmas da escola
            
            n_turmas = variaveis.get_nturmas(key)
            alunos_nee = qt_nees(value)
                                 
            if alunos_nee == 0:
                 novo_valor = N_ALUNOS_TURMA * n_turmas
            else:
                if alunos_nee%2:
                    alunos_nee = alunos_nee + 1
                  
                n_turmas_red = alunos_nee/2
                if n_turmas_red > n_turmas:
                    n_turmas_red = n_turmas
                    
                n_turmas_n_red = n_turmas -  n_turmas_red
                
                if n_turmas_red < 0:
                    n_turmas_red = 0
                elif n_turmas_n_red < 0:
                    n_turmas_n_red = 0
                
                novo_valor =  int(N_ALUNOS_TURMA_REDUZIDA * n_turmas_red + N_ALUNOS_TURMA * n_turmas_n_red)
               
                 
            y = list(lista_escolas[key])
            y[3] = novo_valor
            lista_escolas[key] = tuple(y)           
           

#Conta o número de Nees que reduzem turma, se for mais de 2 na turma, coloca esses NEEs na opção seguinte
def ha_mais_2_nee(lista_colocados,lista_alunos_extra):
    #Se há mais de 2 nees, coloca na lista dos alunos a irem à 2º opção
    
    for key,value in lista_colocados.items():
        if key != 'fora' and key != 'alunos_condicionais':  
            n_nees_reduz = 0
            for x in value:
                if ordenacao.is_reduzturma(x):
                    n_nees_reduz = n_nees_reduz + 1
                    #Quantos alunos extra cabem?
                    n_turmas = lista_escolas[key][2]
                                    
                    if n_nees_reduz > (2*n_turmas):
                        print("o aluno não cabe ",x)
                        lista_alunos_extra.append(x)
    
    #Remove o aluno da lista de colocados                    
    for aluno in lista_alunos_extra:
        for key,value in lista_colocados.items():         
            y = list(value)     
            for aluno_colocado in y:
                if aluno_colocado == aluno:
                    y.remove(aluno)
            lista_colocados[key] = tuple(y)
            
    return lista_alunos_extra

#Coloca nees nas opções seguintes
def coloca_nees_opcao(lista_alunos_extra,lista_colocados,op):  
     
    for aluno in lista_alunos_extra:
        
        #vai buscar a 2ºopção do alunos e as sucessivas
        if op==2:
            opcao_aluno = aluno[12]
        elif op == 3:
            opcao_aluno = aluno[16]
        elif op == 4:
            opcao_aluno = aluno[20]
        elif op == 5:
            opcao_aluno = aluno[24]
            opcao_aluno = ""
        
        if lista_colocados.get(opcao_aluno) != None:
            #o aluno quer uma escola do agrupamento
                y = list(lista_colocados[opcao_aluno])            
                y.append(aluno)
                lista_colocados[opcao_aluno] = tuple(y)
        else:
            # o aluno quer sair do agrupamento ou já não tem opcoes
            y = list(lista_colocados["fora"])
            y.append(aluno)
            lista_colocados["fora"] = tuple(y)


    lista_colocados = ordenacao.ordena_colocados(lista_colocados)
    
    return lista_colocados
    

#Verifica se há alunos que não cabem na turma    
def ha_alunos_extra(lista_colocados):
    
    lista_alunos_extra = []
    
    #Constroi a lista de alunos extra
    #Apaga os alunos que estão a amis
    for key,value in lista_colocados.items():
        if key != 'fora' and key != 'alunos_condicionais':  
            n_alunos = int(lista_escolas[key][3])
            y = list(value)
            lista_alunos_extra = lista_alunos_extra + y[n_alunos:]
            y = y[:n_alunos]
            lista_colocados[key] = tuple(y)
    
    
    return lista_alunos_extra

def coloca_opcao(lista_alunos_extra,op):
    for aluno in lista_alunos_extra:
        #vai buscar a 2ºopção do alunos e as sucessivas
        if op==2:
            opcao_aluno = aluno[12]
        elif op == 3:
            opcao_aluno = aluno[16]
        elif op == 4:
            opcao_aluno = aluno[20]
        elif op == 5:
            opcao_aluno = aluno[24]
        else: 
            opcao_aluno = ""
        
        if lista_colocados.get(opcao_aluno) != None:
            #o aluno quer uma escola do agrupamento
                y = list(lista_colocados[opcao_aluno])            
                y.append(aluno)
                lista_colocados[opcao_aluno] = tuple(y)
                
        else:
            # o aluno quer sair do agrupamento ou já não tem opcoes
            y = list(lista_colocados["fora"])
            y.append(aluno)
            lista_colocados["fora"] = tuple(y)
            
        
      
#Função principal que coloca dos alunos
def coloca_alunos(lista_alunos):
    
    #Coloca todos os alunos na 1 opção
    lista_colocados = coloca_1opcao(lista_alunos) 
    #Ordena os alunos
    lista_colocados = ordenacao.ordena_colocados(lista_colocados)
    print("Ordenamos os alunos na 1ºopção")
    
    print("Vamos tratar os alunos NEE")
    lista_alunos_extra = []
    #Verifica se alguma turma tem mais de 2 NEES, se tiver coloca os NEES nas opções seguintes ou fora
    for i in range(2,6):
        lista_alunos_extra = ha_mais_2_nee(lista_colocados,lista_alunos_extra)
        if lista_alunos_extra != []:
            print("Há alunos NEEs que não tiveram vaga opção ",i-1)
            lista_colocados = coloca_nees_opcao(lista_alunos_extra,lista_colocados,i)
            lista_alunos_extra = []
        else: 
            i = 6
    
    print("Colocamos os alunos NEEs nas opções seguintes")
    #Após colocar NEES, ordena a lista
    print("Ordenamos os alunos com os NEES nas escolas, obedecendo a 2 no máximo por turma")
    lista_colocados = ordenacao.ordena_colocados(lista_colocados)
    
    #De acordo com o NEES vai alterar o número de alunos das turmas
    print("De acordo com o NEES vai alterar o número de alunos das turmas")
    altera_alunos_por_escola(lista_colocados)
    
    #Verifica se há alunos extra e coloca nas opções seguintes os alunos
    print("Verifica se há alunos que não cabem na turma")
    lista_alunos_extra = ha_alunos_extra(lista_colocados)   
    op = 2
    while (lista_alunos_extra != []):
        coloca_opcao(lista_alunos_extra,op)
        print("Coloca esses alunos na opção ",op)
        lista_colocados = ordenacao.ordena_colocados(lista_colocados)
        print("Volta a ordenar ",op)
        lista_alunos_extra = ha_alunos_extra(lista_colocados)   
        print("Verifica se há alunos que não cabem na turma")
        op = op + 1
            

    print("Vou criar a excel com os alunos distribuídos no ficheiro Distribuicao_Final.xlsx")
    processamento_excel.escreve_excel_final(lista_colocados)  


def menu():
    
    print ("Vamos iniciar a constituição das turmas do ficheiro exportacao_portal_seriacao.xlsx")
    lista_alunos = dados_excel.dados_alunos()
    print ("Vamos ler os dados das escolas do ficheiro escolas.xlsx")
    lista_escolas = dados_excel.dados_escolas()
    variaveis.put_escolas(lista_escolas)
    print ("Vamos ler os dados das escolas do ficheiro IPSS.xlsx")
    lista_ipss = dados_excel.dados_ipss()
    variaveis.put_ipss(lista_ipss)
    print ("Vamos criar um ficheiro de excel, com os dados essenciais ListaAlunosMatriculas.xlsx ")
    nome_folha = "ListaAlunosMatriculas.xlsx"    
    processamento_excel.write_result(lista_alunos,nome_folha)
    print ("Vamos colocar os alunos e ordenar")
    coloca_alunos(lista_alunos)
    
   
menu()