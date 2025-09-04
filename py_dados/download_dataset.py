# Run if not installed
# pip install -U pyfabricops
# python.exe -m pip install --upgrade pip para fazer upgrade do pyfabricops

import pyfabricops as pf

Workspace_Name = 'TESTE_DATASIDE'      #Altere o nome do Workspace
Dataset_Name = 'ExportsMinerva_teste'  #Altere o nome do Dataset
Report_Name = 'ExportsMinerva_teste'   #Altere o nome do Relatório
Salvar_Caminho = 'D:/Export_PowerBI/' + Dataset_Name

pf.set_auth_provider('oauth')
pf.setup_logging(format_style='minimal') 

# Baixando o modelo semântico
pf.export_semantic_model(
    Workspace_Name,
    Dataset_Name,
    Salvar_Caminho,
)

# Baixando o relatório
pf.export_report(
    Workspace_Name,
    Report_Name,
    Salvar_Caminho,
)

# Reapontamento do relatório para o modelo semântico local
pf.convert_report_definition_to_by_path(
    Salvar_Caminho + Report_Name,
    Salvar_Caminho,
)