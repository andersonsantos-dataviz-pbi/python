# Run if not installed
# pip install -U pyfabricops
# python.exe -m pip install --upgrade pip para fazer upgrade do pyfabricops

import pyfabricops as pf

#Declarando variáveis
Workspace_Name = 'BI Store'      #Altere o nome do Workspace
Dataset_Name = 'Vivo Pré'  #Altere o nome do Dataset

#Autenticando através do web browse
pf.set_auth_provider('oauth')
pf.setup_logging(format_style='minimal')

#Efetuando a atualização de partições do modelo semântico
pf.refresh_semantic_model(
    Workspace_Name,
    Dataset_Name,
    apply_refresh_policy=False, #Desabilita a politica de atualização da incremental
#Caso deseje atualizar mais de uma tabela e mais de uma partição, basta colocar (,) após a última chave e acrescentar
#mais um conjunto de chaves com a próxima tabela e partição a ser atualizada
    objects=[
        {
            "table": "resgate" ,
            "partition": "2025Q307"
        }
    ]
)