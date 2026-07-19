# TRACKER

Sistema de gerenciamento de projetos e tarefas desenvolvido para uso interno, com integração ao Microsoft SharePoint para armazenamento de seeds (snapshots de dados).

## Visão Geral

O TRACKER é uma aplicação web single-page (Flask + HTML/CSS/JS vanilla) que permite:

- Gerenciar projetos e demandas por status em visões Kanban, Colunas e Fluxograma
- Controlar usuários com sistema de permissões (admin / membro)
- Registrar solicitantes e fontes de informação via cadastro de contatos
- Filtrar demandas por especialista responsável
- Salvar e restaurar snapshots (seeds) do estado do sistema no SharePoint
- Gerar apresentações PowerPoint por projeto e status
- Acompanhar um Dashboard com resumo por especialista e demandas críticas

A autenticação é feita via **Microsoft Device Code Flow** (MSAL), sem necessidade de registrar um app próprio no Azure AD.

---

## Pré-requisitos

- Python 3.10+
- Conta Microsoft com acesso ao SharePoint da organização
- Windows (o armazenamento seguro do token usa DPAPI — Windows Data Protection API)

---

## Instalação

```bash
# 1. Clone o repositório
git clone https://github.com/seu-usuario/tracker.git
cd tracker

# 2. Crie o ambiente virtual
python -m venv .venv
.venv\Scripts\activate   # Windows
# source .venv/bin/activate  # Linux/Mac (sem suporte a DPAPI)

# 3. Instale as dependências
pip install -r requirements.txt

# 4. Configure
cp config.example.py config.py
# Edite config.py com os valores do seu ambiente (SharePoint, porta, etc.)
```

---

## Configuração

Edite o arquivo `config.py` (criado a partir do `config.example.py`):

| Variável | Descrição |
|---|---|
| `AAD_CLIENT_ID` | Client ID do app Azure AD (use o público do Office 365 ou registre o seu) |
| `AAD_AUTHORITY` | URL de autoridade (`common` para multi-tenant) |
| `SCOPES` | Escopos Microsoft Graph necessários |
| `SP_HOST` | Domínio do SharePoint (ex: `suaempresa.sharepoint.com`) |
| `SP_SITE` | Nome do site SharePoint |
| `SP_FOLDER` | Pasta raiz do TRACKER no SharePoint |
| `SP_SEEDS` | Subpasta para seeds dentro de `SP_FOLDER` |
| `PORT` | Porta local do servidor (padrão: `5003`) |
| `HOST` | Host do servidor (padrão: `0.0.0.0`) |

> **Importante:** `config.py` está no `.gitignore` e **nunca deve ser commitado**.

---

## Uso

```bash
.venv\Scripts\activate
python app.py
```

O navegador abre automaticamente. No primeiro acesso, o sistema inicia o **Device Code Flow**:

1. Um código é exibido na tela
2. Acesse [microsoft.com/devicelogin](https://microsoft.com/devicelogin)
3. Digite o código e faça login com sua conta Microsoft
4. O sistema autentica e carrega os dados do SharePoint

---

## Estrutura do Projeto

```
tracker/
├── app.py                 # Backend Flask — rotas, lógica, integração SharePoint
├── config.py              # Configuração local (NÃO commitado)
├── config.example.py      # Template de configuração para novos ambientes
├── requirements.txt       # Dependências Python
├── .gitignore
├── seeds/                 # Seeds locais (backup offline)
│   └── .gitkeep
└── templates/
    ├── index.html         # SPA principal
    ├── auth.html          # Tela de autenticação Device Code
    └── config.html        # Painel de configurações
```

---

## Seeds

Seeds são snapshots completos do estado do sistema (projetos, itens, contatos) salvos no SharePoint.

- **Salvar seed**: aba *Seeds → Criar / Fechar*
- **Restaurar seed**: aba *Seeds → Seeds Salvos → ↩ Restaurar*
- **Atualizar lista do SharePoint**: botão *↺ Atualizar do SharePoint* (a lista fica em cache até ser atualizada manualmente)

---

## Rate Limiting

Rotas sensíveis possuem limite de requisições por IP:

| Rota | Limite |
|---|---|
| `/auth` | 10 / minuto |
| `POST /api/seeds` | 10 / hora |
| `POST /api/seeds/refresh` | 5 / minuto |

---

## Segurança

- O token de acesso Microsoft é armazenado localmente em `.sp_token_cache.bin`, criptografado com **DPAPI** (vinculado ao usuário e máquina Windows).
- Nenhuma credencial é armazenada em texto plano.
- Apenas o proprietário de um item ou um administrador pode editá-lo ou excluí-lo.
- `config.py`, `tracker_data.json` e `.sp_token_cache.bin` estão no `.gitignore`.

---

## Dependências

| Pacote | Uso |
|---|---|
| `flask` | Framework web |
| `flask-limiter` | Rate limiting |
| `msal` | Autenticação Microsoft (Device Code Flow) |
| `requests` | Chamadas à Microsoft Graph API |
| `python-pptx` | Geração de apresentações PowerPoint |

---

