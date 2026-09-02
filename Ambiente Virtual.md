Criar e configurar um ambiente virtual (venv) no VS Code é um processo rápido. Para isso, abra a pasta do projeto no editor, abra o terminal integrado, gere o ambiente com o comando python -m venv .venv e selecione-o na paleta de comandos. https://www.youtube.com/watch?v=wOchmO8J7gA\&t=21



**Passo a Passo para Criar e Ativar a Venv**

1. **Abra o seu projeto:** Inicie o VS Code e abra a pasta raiz do seu projeto. https://www.youtube.com/watch?v=wOchmO8J7gA\&t=21
2. **Abra o terminal integrado:** Use o atalho Ctrl + ' (ou navegue pelo menu superior em Terminal > New Terminal). https://dev.to/franciscojdsjr/guia-completo-para-usar-o-virtual-environment-venv-no-python-57bo
3. **Crie a ambiente virtual:** Execute o comando abaixo no terminal do VS Code. Este comando criará uma pasta oculta chamada .venv contendo a cópia local do Python:
**bash**
python -m venv .venv
4. **Ative a venv:** O comando de ativação depende do seu sistema operacional:

   * Windows (Command Prompt): 
**cmd**
.venv\\Scripts\\activate.bat
   * Windows (PowerShell):
**powershell**
.venv\\Scripts\\Activate.ps1



**Configurando no VS Code**

Após a ativação, o nome do ambiente virtual aparecerá entre parênteses no início da linha do seu terminal (ex: (.venv) C:\\SeuProjeto>). Agora, informe ao VS Code qual interpretador ele deve usar: https://cursos.alura.com.br/forum/topico-vscode-ambiente-virtual-python-347100

1. Pressione Ctrl + Shift + P (ou Cmd + Shift + P no Mac) para abrir a Paleta de Comandos.
2. Digite e selecione: Python: Select Interpreter.
3. Escolha a opção correspondente ao seu ambiente recém-criado (geralmente indicada pelo caminho .venv ou Python 3.x.x 64-bit ('.venv': venv)). https://cursos.alura.com.br/forum/topico-vscode-ambiente-virtual-python-347100

Para acompanhar o processo visualmente e conferir como o editor gerencia os interpretadores e dependências: https://www.youtube.com/watch?v=wOchmO8J7gA



**Dicas Úteis**



* **Instalando pacotes:** Com a venv ativa, qualquer pacote instalado via pip install <nome-do-pacote> ficará isolado neste projeto.
* **Salvando dependências:** Para salvar os pacotes instalados em um arquivo que pode ser compartilhado, utilize pip freeze > requirements.txt.
* **Documentação oficial:** Para mais detalhes e configurações avançadas de ambiente, consulte a documentação oficial de Python environments in VS Code.

