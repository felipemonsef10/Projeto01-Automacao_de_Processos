# 📊 Projeto 01 – Automação de Indicadores

Este projeto foi desenvolvido como parte de um curso de Python voltado para automação de tarefas. O objetivo foi simular o dia a dia de um analista de dados em uma rede varejista, criando um sistema automatizado de envio de relatórios diários e mensais (OnePages) para os gerentes de loja e para a diretoria.

## 🚀 Objetivo

Automatizar o processo de:
- Cálculo dos principais indicadores de desempenho por loja.
- Geração e envio de relatórios personalizados por e-mail.
- Organização e salvamento dos arquivos por loja e por data.
- Envio de resumo consolidado para a diretoria com rankings de desempenho.

## 📂 Estrutura dos Arquivos

- `Emails.xlsx` → Lista de e-mails de gerentes e diretoria.
- `Vendas.xlsx` → Base de dados com todas as vendas realizadas.
- `Lojas.csv` → Lista com o nome das lojas.
- `onepage.png` e `Exemplo.JPG` → Exemplos visuais de como os relatórios são gerados e enviados.

## 📈 Indicadores Calculados

- **Faturamento**
- **Diversidade de Produtos Vendidos**
- **Ticket Médio por Venda**

Cada indicador é calculado com base:
- No **dia mais recente** da base de vendas.
- No acumulado **do ano** até a data mais recente.

## 🔧 Tecnologias e Bibliotecas Utilizadas

- `Python`
- `pandas` → Manipulação e análise de dados.
- `smtplib` / `email` → Envio automatizado de e-mails com corpo HTML e anexos.
- `openpyxl` → Leitura e escrita de arquivos Excel.
- `os`, `datetime`, `pathlib` → Automação de diretórios e manipulação de datas.

## 🧠 Conceitos Aplicados

- Automação de tarefas com Python
- Manipulação de planilhas Excel
- Envio de e-mails personalizados com anexos
- Criação de backups organizados por data
- Geração de rankings com base em métricas

## 📨 Funcionalidades

✔️ Gera relatórios personalizados por loja  
✔️ Envia e-mails com o relatório no corpo e dados em anexo  
✔️ Cria ranking de desempenho diário e anual para a diretoria  
✔️ Armazena o histórico de relatórios por loja  

## 📬 Exemplo de Envio

Email para gerente de loja:
- Corpo do e-mail com resumo (OnePage)
- Anexo com os dados de vendas da respectiva loja

Email para diretoria:
- Destaque para a melhor e pior loja (dia e ano)
- Anexos com rankings diários e anuais

## 📌 Como Executar

1. Instale os requisitos:
```bash
pip install pandas openpyxl
````

2. Atualize os e-mails no arquivo `Emails.xlsx` com seus próprios endereços para testes.

3. Execute o script principal com:

```bash
python seu_script.py
```

⚠️ Certifique-se de ativar o acesso de apps menos seguros em sua conta de e-mail (se estiver usando Gmail) ou configurar uma senha de app.

## 🔗 Autor

Desenvolvido por **Felipe Monsef**
[LinkedIn](https://www.linkedin.com/in/felipe-monsef-71510b247/) • [GitHub](https://github.com/felipemonsef10)

---

> Este projeto é uma simulação educacional e não utiliza dados reais.
