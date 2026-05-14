# Contagem de Estoque 📦

Sistema de contagem de estoque com suporte a múltiplos patrimônios por produto, inventários e exportação de relatórios.

## Como usar

1. Acesse o app pelo link do Streamlit Cloud
2. Crie um inventário na barra lateral
3. Carregue o arquivo `estoque.xlsx` pelo botão **📂 Carregar / Atualizar Estoque**
4. Informe seu nome no campo **Operador**
5. Use a aba **🔍 Contar Item** para registrar as contagens

## Deploy no Streamlit Cloud (com banco de dados persistente)

O app detecta automaticamente se deve usar **SQLite** (local) ou **PostgreSQL via Supabase** (nuvem).  
Para que os dados persistam no deploy, configure o Supabase:

### 1. Criar banco no Supabase

1. Acesse [supabase.com](https://supabase.com) e crie um projeto gratuito
2. No painel do projeto vá em **Project Settings → Database → Connection string → URI (Transaction pooler)**
3. Copie a URI no formato:  
   `postgresql://postgres.<REF>:<SENHA>@aws-0-<REGION>.pooler.supabase.com:6543/postgres`

### 2. Configurar secrets no Streamlit Cloud

1. No painel do seu app em [share.streamlit.io](https://share.streamlit.io), acesse **Settings → Secrets**
2. Cole o conteúdo abaixo com a URI real:

```toml
SUPABASE_URL = "postgresql://postgres.<REF>:<SENHA>@aws-0-<REGION>.pooler.supabase.com:6543/postgres"
```

> O arquivo `.streamlit/secrets.toml.example` serve de referência — **nunca commite o `secrets.toml` com credenciais reais**.

### 3. Publicar no GitHub e fazer o deploy

```bash
git init
git add .
git commit -m "first commit"
git remote add origin https://github.com/<seu-usuario>/<seu-repo>.git
git push -u origin main
```

4. Acesse [share.streamlit.io](https://share.streamlit.io) → **New app** → selecione o repositório e `contagem.py` → **Deploy**

## Rodando localmente

```bash
pip install -r requirements.txt
# Preencha .streamlit/secrets.toml com base no secrets.toml.example (opcional — sem ele usa SQLite)
streamlit run contagem.py
```

## Observações

- **Sem `SUPABASE_URL`**: o app usa SQLite local (`contagem.db`), criado automaticamente
- **Com `SUPABASE_URL`**: usa PostgreSQL no Supabase — dados persistem entre deploys e reinicializações
- O arquivo `estoque.xlsx` pode ser carregado diretamente pelo app, sem acesso ao servidor
