
lista_ipss = {
    "803282":["Fundação Joquim Oliveira Lopes - Pólo I","AVINTES"],
    "521700":["Jardim Infantil Da Fundação Padre Luís", "OLIVEIRA DO DOURO"],
    "804089":["Jardim de Infância do Centro Salvador Caetano e Ana Caetano", "VILAR DE ANDORINHO"],
    "803283":["Fundação Joquim Oliveira Lopes - Pólo II","AVINTES"]
}    

lista_escolas = {
        "247560":("Escola Básica de Freixieiro","OLIVEIRA DO DOURO",2,48),
        "273946":("Escola Básica de Sardão", "OLIVEIRA DO DOURO",1,24),
        "284646":("Escola Básica de Vilar", "VILAR DE ANDORINHO", 1,24),
        "297148":("Escola Básica Dr. Fernando Guedes","AVINTES",2,48),
        "237279":("Escola Básica de Cabanões", "AVINTES",1,24),
        "201844":("Escola Básica de Aldeia Nova", "AVINTES",1,24),
        "231514":("Escola Básica de Mariz","VILAR DE ANDORINHO",0,0),
        "fora": ("Fora do agrupamento"),
        "alunos_condicionais":("Alunos Condicionais")
}

lista_colocados = {
        "247560":(""),
        "273946":(""),
        "284646":(""),
        "297148":(""),
        "237279":(""),
        "201844":(""),
        "231514":(""),
        "fora":(""),
        "alunos_condicionais":("")    
}

op = 1

def get_nturmas(key):
    return lista_escolas[key][2]
    
def get_listaColocados():
    return lista_colocados

def get_op():
    return op

def get_escolas():
    return lista_escolas

def get_ipss():
    return lista_ipss

def put_ipss(nova_lista):
    global lista_ipss      
    lista_ipss = nova_lista
    for key,value in lista_ipss.items():
        print("A ipss "+key+" - "+value[0]+" situa-se em "+value[1])
       
def put_escolas(nova_lista):
    global lista_escolas
    lista_escolas = nova_lista
    for key,value in lista_escolas.items():
        print("A escola "+key+" - "+value[0]+" tem "+str(value[2])+ " turmas")