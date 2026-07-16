# ── TRACKER — Exemplo de Configuração ────────────────────────────────────────
# Copie este arquivo para config.py e preencha com os valores reais.
# NUNCA commite config.py no Git.

# ── Microsoft / Azure AD ──────────────────────────────────────────────────────
# Para usar o app público do Office 365 (mesma abordagem do Torre Web):
AAD_CLIENT_ID = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'
AAD_AUTHORITY = 'https://login.microsoftonline.com/common'
SCOPES        = ['https://graph.microsoft.com/Files.ReadWrite.All']

# Ou registre seu próprio app no Azure AD e use:
# AAD_CLIENT_ID = 'seu-client-id-aqui'
# AAD_AUTHORITY = 'https://login.microsoftonline.com/seu-tenant-id'

# ── SharePoint ────────────────────────────────────────────────────────────────
SP_HOST   = 'seudominio.sharepoint.com'
SP_SITE   = 'NOME_DO_SITE'
SP_FOLDER = 'NOME_DA_PASTA'
SP_SEEDS  = 'seeds'

# ── Servidor ──────────────────────────────────────────────────────────────────
PORT = 5003
HOST = '0.0.0.0'
