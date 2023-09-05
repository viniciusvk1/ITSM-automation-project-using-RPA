<h1>Automação RPA - ITSM SoftExpert (Cancelamento de chamados)</h1>

<h2> Documentação de Instalação e Uso do Código </h2>

Este documento fornece instruções sobre como instalar e usar o código Python fornecido para automatizar ações usando a biblioteca `pyautogui` e processar dados com o `pandas`. O código em questão lê um arquivo Excel, interage com uma aplicação na tela do computador e realiza ações com base nos dados lidos. Certifique-se de ter as bibliotecas Python necessárias instaladas antes de executar o código.
 
<h2>Pré-requisitos</h2>

* Python 3.x (Recomendado)
* pandas
* pyautogui

<h3>Você pode instalar essas bibliotecas usando o pip: </h3>

~~~bash
pip install PACKAGENAME pyautogui
~~~

OU

~~~bash
python -m pip install PACKAGENAME --trusted-host=pypi.python.org --trusted-host=pypi.org --trusted-host=files.pythonhosted.org
~~~

<h2>Instruções de Uso</h2>

1. Preparação do Ambiente:

Certifique-se de que você tenha o Python 3.x instalado e tenha instalado as bibliotecas pandas e pyautogui usando o comando pip conforme mencionado acima.

2. Configuração do Excel:

Antes de executar o código, verifique se o caminho do arquivo Excel e o nome da coluna estão corretos. Abra o código e atualize as seguintes variáveis de acordo com suas necessidades:
<br> <br>`excel_file_path`: O caminho completo para o arquivo Excel que você deseja ler.
<br> `column_name`: O nome da coluna no arquivo Excel que contém os dados que serão usados no script.

3. Personalização:

Personalize as ações e o comportamento do código de acordo com suas necessidades, ajustando as coordenadas, tempos de espera e ações específicas realizadas no aplicativo alvo.

Lembre-se de que a automação de interface do usuário pode ser sensível às mudanças na interface ou nas resoluções de tela, portanto, esteja preparado para fazer ajustes no código conforme necessário.

<h1>Avisos e Considerações</h1>

* Use este código com responsabilidade e apenas em ambientes e aplicativos que você tem permissão para automatizar.
* Certifique-se de que a aplicação alvo esteja em um estado consistente e previsível para evitar erros na automação.
* Este código pode requerer ajustes e testes adicionais, dependendo do ambiente em que você o utiliza.

**Lembre-se de fazer backup de seus dados e código antes de automatizar tarefas críticas, e teste-o em um ambiente seguro antes de usá-lo em produção.**