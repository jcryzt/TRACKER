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
SP_HOST   = 'seudominio.sharepoint.com'   # ex: contoso.sharepoint.com
SP_SITE   = 'NOME_DO_SITE'               # nome do site SharePoint (sem /sites/)
SP_FOLDER = 'NOME_DA_PASTA'              # pasta raiz onde fica tracker_data.json

# Subpasta dentro de SP_FOLDER onde os seeds sao armazenados no SharePoint.
# O Tracker faz upload/download automatico desta pa