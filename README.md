# ⚡ EQUATORIAL CYBORG
### Robô de Download de Faturas + Gerador Profissional de Relatórios Excel

Sistema desenvolvido para automatizar:

- 🤖 Download de faturas no site da Equatorial
- 📂 Organização automática de PDFs
- 📊 Extração de dados técnicos e financeiros
- 📈 Geração de relatório Excel profissional e formatado

---

# 📌 Visão Geral

O projeto é dividido em dois módulos principais:

## 1️⃣ Robô de Automação Web
Responsável por:
- Login manual assistido
- Navegação entre contas contrato (UC)
- Download automático da última fatura disponível
- Organização automática dos arquivos PDF

## 2️⃣ Gerador de Relatórios
Responsável por:
- Leitura de todos os PDFs baixados
- Extração de dados técnicos, tributários e financeiros
- Aplicação da regra de competência por data de leitura
- Geração de relatório Excel estruturado e formatado

---

# 📁 Estrutura do Projeto

```bash
EQUATORIAL_AUTOMACAO/
│
├── login.bat # Inicializa o robô
├── executar.bat # Executa o gerador de relatórios
├── requirements.txt
│
├── src/
│ ├── app_hibrido.py
│ ├── assistente_login.py
│ ├── extrator.py
│ ├── gerador_faturas.py
│ ├── leitor_credenciais.py
│ ├── organizador_visual.py
│ └── main.py
│
├── output/
│ ├── faturas/ # PDFs baixados
│ ├── relatorios/ # Excel final gerado
│ └── debug/
│
└── perfil_bot/
```


---

# 🚀 Como Executar

## ▶️ 1. Executar o Robô

Dê duplo clique em:

```bash
login.bat
```

O sistema irá:
- Verificar ambiente Python
- Iniciar painel de controle
- Abrir o navegador automaticamente

⚠️ Não feche o terminal durante a execução.

---

## 🔐 2. Login

O login deve ser feito manualmente no site.

Após estar logado, utilize o botão do painel:



ROBÔ BAIXAR ÚLTIMA FATURA


O sistema irá:
- Baixar o PDF
- Salvar em `output/faturas`
- Fazer logout automático (em caso de sucesso)

---

## 📊 3. Gerar Relatório Excel

Após baixar todas as faturas:


```bash
executar.bat
```

Escolha:



[1] Criar relatório profissional


Informe o mês de referência (ex: `02/2026`).

O sistema irá:
- Ler todos os PDFs
- Extrair os dados
- Aplicar regra de competência
- Gerar o Excel final em `output/relatorios`

---

# 🧠 Regra de Competência

A competência contábil é definida pela data de leitura:

- 📅 Leitura até dia 12 → Conta como mês atual  
- 📅 Leitura após dia 12 → Conta como mês seguinte  

Essa regra é aplicada automaticamente.

---

# 📊 Estrutura do Relatório

O Excel gerado contém:

- Aba **DETALHES**
- Aba **RESUMO**
- Aba **ESTATÍSTICAS**

Os dados são organizados em grupos:

- 🔵 Identificação
- 🟠 Datas
- 🟢 Medição
- 🔴 Valores financeiros
- 🟣 Preços unitários
- 🔵 Tributos
- 🟤 Informações técnicas
- ⚙️ Controle de extração

Caso algum campo não seja encontrado no PDF:
- O sistema preenche com `0.00`
- Ou registra na coluna **Erro Extração**

---

# 🏗️ Como Foi Desenvolvido

## 🔹 Automação Web
- Selenium
- Undetected ChromeDriver
- WebDriver Manager

O robô possui:
- Detecção de múltiplas UCs
- Tratamento de troca de conta contrato
- Tentativas automáticas em caso de falha
- Fechamento automático de assistentes virtuais

---

## 🔹 Extração de Dados

O módulo de processamento:

- Lê os PDFs baixados
- Aplica expressões regulares (regex)
- Normaliza datas e valores
- Realiza cruzamento com base de dados interna
- Gera planilha formatada com XlsxWriter

---

## 🔹 Geração do Excel

Bibliotecas utilizadas:

- pandas
- openpyxl
- xlsxwriter

Recursos aplicados:

- Formatação por cores por grupo
- Ajuste automático de colunas
- Formatação monetária
- Cálculos automáticos
- Separação em múltiplas abas

---

# 📦 Instalação

Instale as dependências:


```bash
pip install -r requirements.txt
```

Ou manualmente:


```bash
pip install pandas openpyxl xlsxwriter selenium webdriver-manager undetected-chromedriver pyperclip
```

---

# 🖥️ Requisitos

- Python 3.10+
- Google Chrome instalado
- Windows 10 ou superior
- Conexão com internet

---

# ⚠️ Problemas Conhecidos

### Página não responde na primeira execução
Pressione `F5` no navegador e tente novamente.

### Troca automática de UC pode falhar
Clique em **Tentar novamente**.

### Mudanças no layout do site
O robô pode precisar de atualização.

---

# 🛠️ Suporte

Em caso de erro crítico:

1. Tire print da tela do terminal
2. Informe qual módulo estava executando
3. Envie ao desenvolvedor

---

# 👨‍💻 Autor

**Gabriel Teperino**  
Automação • Python • Extração de Dados • Relatórios Excel Profissionais
