# Conectar ao Power BI Desktop (local)

Este pequeno projeto contém um script que detecta a porta do Analysis Services do Power BI Desktop em execução e tenta se conectar ao modelo via `pyadomd` para executar uma consulta DAX de teste.

Passos rápidos:

1. Abra o Power BI Desktop com o relatório desejado.
2. No terminal (PowerShell):

```powershell
python -m pip install -r requirements.txt
python connect_powerbi.py
```

Observações:

- O `pyadomd` requer o driver ADOMD.NET presente no Windows (normalmente instalado com ferramentas cliente do Analysis Services ou Visual Studio). Se receber erro relacionado ao provedor, instale o componente ADOMD.NET da Microsoft.
- O script procura pelo processo `msmdsrv.exe` (instância do Analysis Services usada pelo Power BI Desktop) e obtém a porta TCP onde ele está ouvindo.
- A consulta de teste usada é `EVALUATE ROW("Ping", 1)` — altere conforme necessário.
