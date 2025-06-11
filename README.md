# Automação de Relatórios e Envio – Excel com VBA, Power BI e Python utilizando as bibliotecas selenium, pyautogui, win32, pandas e openpyxl.

Este projeto automatiza o download de relatórios em Excel via Selenium, processa e organiza os dados com VBA, gera visualizações no Power BI, tira screenshots dos dashboards e envia por e-mail os arquivos finais utilizando Python.

---

## 🛠️ Tecnologias Utilizadas

- **Python** (selenium, pyautogui, win32, pandas e openpyxl, pymsgbox)
- **Excel VBA**
- **Power BI**
- **Git** (para versionamento)

---

## ⚙️ Fluxo da Automação

1. **Download Automático**
   - A automação acessa o sistema via Selenium e faz o download do relatório Excel.

2. **Processamento dos Dados**
   - Um script VBA organiza, limpa e atualiza os dados na base.

3. **Atualização de dashboards**
   - O Power BI é atualizado automaticamente com a biblioteca pyautogui, simulando os cliques na tela.
   - Um script Python tira print do dashboard atualizado.

4. **Envio de E-mail**
   - O Python envia o Excel atualizado e o print do Power BI no corpo do Email para os destinatários via Outlook.
