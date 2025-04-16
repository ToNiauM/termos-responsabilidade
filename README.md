# Termos de Responsabilidade - CFC

Este sistema foi desenvolvido para gerar documentos oficiais de Termo de Responsabilidade e Termo de Devolução de bens patrimoniais, conforme as diretrizes do Conselho Federal de Contabilidade (CFC).

## 📁 Estrutura do Projeto

| Arquivo | Função |
|--------|--------|
| `app.py` | Roteador principal da aplicação Flask, responsável por servir as páginas web e acionar os scripts de geração de documentos. |
| `Script_Termo_Individual.py` | Gera o termo de responsabilidade para **um servidor específico**, com base no nome e nos bens associados. |
| `Termo_de_Responsabilidade.py` | Gera o termo por **centro de custo (unidade organizacional)**, agrupando todos os bens e salvando o documento `.docx` + planilha. |
| `termo_devolucao.py` | Gera o termo de devolução para bens previamente cadastrados na planilha `geral.xlsx`, com base nos números informados. |
| `timbrado.docx` | Modelo do documento Word com o papel timbrado oficial. |
| `acervo.xlsx` / `geral.xlsx` | Planilhas de entrada contendo os dados de bens e responsáveis. |
| `requirements.txt` | Lista de bibliotecas necessárias para rodar o sistema. |
| `Procfile` | Arquivo que instrui o Render a iniciar a aplicação usando `gunicorn`. |

---

## ▶️ Como Executar Localmente

1. **Crie um ambiente virtual (recomendado):**

```bash
python -m venv venv
venv\Scripts\activate  # Windows
