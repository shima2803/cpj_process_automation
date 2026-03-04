# CPJ Process Automation

Automação desenvolvida em Python para realizar o download de documentos de processos no sistema **CPJ** de forma automatizada.  
O script lê uma lista de **Litigation IDs** a partir de uma planilha Excel, acessa cada processo dentro do sistema, realiza o download dos documentos e organiza os arquivos automaticamente em pastas separadas.

Essa ferramenta foi criada para **reduzir trabalho manual repetitivo**, agilizando a coleta de documentos em ambientes jurídicos que lidam com grande volume de processos.

---

# Como funciona

O fluxo do script segue as seguintes etapas:

1. O programa inicia e abre o sistema **CPJ**.
2. O login é realizado automaticamente utilizando credenciais configuradas via variáveis de ambiente.
3. O script lê uma planilha Excel contendo os **Litigation IDs**.
4. Para cada ID encontrado:
   - O processo é aberto no CPJ.
   - O menu de documentos é acessado.
   - O documento é salvo utilizando automação de interface.
5. O download é monitorado na pasta **Downloads** do sistema.
6. O arquivo é movido automaticamente para uma pasta específica do processo.
7. O script retorna para a tela inicial e continua com o próximo processo.
8. Ao finalizar todos os processos, o programa encerra o CPJ.

---

# Estrutura de pastas

A estrutura básica do projeto é a seguinte:


```
project/
│
├── main.py
├── input.xlsx
├── assets/
│ └── save.png
└── documents_by_litigation_id/
``` 


## Descrição das pastas

### main.py
Script principal responsável por executar toda a automação.

### input.xlsx
Planilha contendo a lista de **Litigation IDs** que serão processados.

### assets/
Pasta que contém imagens utilizadas pelo PyAutoGUI para localizar elementos na tela.

### documents_by_litigation_id/
Pasta criada automaticamente onde os documentos baixados serão organizados.

Cada processo terá sua própria subpasta.

Exemplo:

```
documents_by_litigation_id/
├── 123456
│ └── documento.pdf
├── 123457
│ └── documento.pdf
```


---

# Requisitos

Para executar o projeto é necessário ter:

## 1. Python instalado

Versão recomendada:

Python 3.10+


---

## 2. Instalar as bibliotecas necessárias

Execute o seguinte comando:

pip install pandas pyautogui openpyxl


Bibliotecas utilizadas:

- **pandas** → leitura da planilha Excel
- **openpyxl** → engine para arquivos Excel
- **pyautogui** → automação da interface do sistema
- **pathlib / os / subprocess** → manipulação de arquivos e execução de processos

---

# Configuração

Algumas configurações são feitas por **variáveis de ambiente** para evitar exposição de dados sensíveis.

## Variáveis necessárias

```
CPJ_EXECUTABLE_PATH
CPJ_USERNAME
CPJ_PASSWORD
```

### Exemplo

```
CPJ_EXECUTABLE_PATH=C:\path\to\cpj3cclient.exe
CPJ_USERNAME=seu_usuario
CPJ_PASSWORD=sua_senha
```
---

# Formato da planilha Excel

A planilha deve conter uma coluna chamada:


LitigationID


Exemplo:

| LitigationID |
|---------------|
| 123456 |
| 123457 |
| 123458 |

O script irá ler todos os valores dessa coluna e processar cada um automaticamente.

---

# Como executar

Depois de configurar tudo, execute:


python main.py


O script irá:

1. Abrir o sistema CPJ
2. Fazer login
3. Ler a planilha
4. Processar cada processo
5. Baixar os documentos
6. Organizar os arquivos automaticamente

---

# Observações

- O script utiliza **automação de interface**, portanto o computador não deve ser utilizado durante a execução.
- A resolução de tela pode influenciar na detecção de elementos.
- O sistema CPJ precisa estar acessível e funcionando corretamente.

---

# Objetivo do projeto

Este projeto foi criado para demonstrar **automação de tarefas repetitivas em sistemas desktop**, utilizando Python e automação de interface.

Principais objetivos:

- Automatizar download de documentos
- Reduzir trabalho manual
- Processar grandes volumes de processos
- Organizar arquivos automaticamente

---

# Licença

Este projeto é apenas para fins educacionais e demonstração de automação.
