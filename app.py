from flask import Flask, jsonify, request, render_template, send_file as sf, redirect
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
import json, uuid, os, sys, subprocess, threading, webbrowser, time, logging
from datetime import datetime, timezone, timedelta

# ── PyInstaller compatibility ─────────────────────────────────────────────────
def _base_dir():
    """Diretório onde ficam os dados (config.py, tracker_data.json, seeds/).
    Quando frozen (.exe): diretório do executável.
    Quando script: diretório do app.py."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def _resource_dir():
    """Diretório onde ficam os recursos empacotados (templates/, static/).
    Quando frozen: sys._MEIPASS (pasta temporária do PyInstaller).
    Quando script: diretório do app.py."""
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

# Garante que config.py ao lado do .exe seja importável
_bd = _base_dir()
if _bd not in sys.path:
    sys.path.insert(0, _bd)

# ── Logging: arquivo ao lado do exe (sem console quando frozen) ───────────────
_log_file = os.path.join(_bd, 'tracker.log')
_handlers = [logging.FileHandler(_log_file, encoding='utf-8')]
if not getattr(sys, 'frozen', False):
    _handlers.append(logging.StreamHandler())
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s — %(message)s",
    handlers=_handlers,
)

def _fatal(msg: str):
    """Exibe mensagem de erro visível mesmo sem console e encerra."""
    logging.critical(msg)
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "Tracker — Erro Fatal", 0x10)
    except Exception:
        pass
    sys.exit(1)

try:
    from config import (
        AAD_CLIENT_ID as _AAD_CLIENT_ID,
        AAD_AUTHORITY as _AAD_AUTHORITY,
        SCOPES        as _SCOPES,
        SP_HOST, SP_SITE, SP_FOLDER, SP_SEEDS,
        PORT, HOST,
    )
except ImportError:
    _fatal(
        "config.py não encontrado.\n\n"
        "Copie config.example.py → config.py e preencha os valores.\n\n"
        f"Caminho esperado:\n{os.path.join(_bd, 'config.py')}"
    )
except Exception as e:
    _fatal(f"Erro ao carregar config.py:\n{e}")

app = Flask(__name__,
    template_folder=os.path.join(_resource_dir(), 'templates'),
    static_folder=os.path.join(_resource_dir(), 'static'),
)
limiter = Limiter(
    key_func=get_remote_address,
    app=app,
    default_limits=[],
    storage_uri="memory://",
)

BASE_DIR  = _base_dir()
DATA_FILE = os.path.join(BASE_DIR, 'tracker_data.json')
SEED_DIR  = os.path.join(BASE_DIR, 'seeds')
os.makedirs(SEED_DIR, exist_ok=True)

# ── CONSTANTS ─────────────────────────────────────────────────────────────────
GRAPH     = 'https://graph.microsoft.com/v1.0'
# _AAD_CLIENT_ID, _AAD_AUTHORITY, _SCOPES, SP_HOST, SP_SITE, SP_FOLDER, SP_SEEDS, PORT, HOST
# are all loaded from config.py above

MESES_PT         = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ']
STATUS_PENDENTES   = ['backlog','andamento','revisao']
STATUS_FINALIZADOS = ['concluido']
STATUS_OUTROS      = ['descartado']

def now(): return datetime.now().strftime('%d/%m/%Y')
def seed_name_auto(especialista=None):
    d = datetime.now()
    mes = MESES_PT[d.month-1]
    ano = str(d.year)[2:]
    if especialista:
        import re as _re
        safe = _re.sub(r'[^A-Z0-9]', '_', especialista.strip().upper())
        safe = _re.sub(r'_+', '_', safe).strip('_')
        return f"SEED_{safe}_{mes}.{ano}"
    return f"SEED_{mes}.{ano}"

def _parse_especialista(filename):
    """SEED_JOAO_JUL.26.json -> 'JOAO' | SEED_JUL.26.json -> None"""
    import re as _re
    name = filename.replace('.json', '')
    parts = name.split('_')
    # formato: SEED_{ESP}_{MES}.{ANO}  — pelo menos 3 partes
    if len(parts) >= 3 and parts[0] == 'SEED':
        # ultima parte deve ser MES.ANO (ex: JUL.26)
        if _re.match(r'^[A-Z]{3}\.\d{2}$', parts[-1]):
            return '_'.join(parts[1:-1])
    return None

# ── DPAPI — copiado do Torre Web ──────────────────────────────────────────────
import ctypes, ctypes.wintypes as _wintypes

class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", _wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

def _dpapi_encrypt(dados: bytes):
    try:
        buf   = ctypes.create_string_buffer(dados)
        b_in  = _DataBlob(len(dados), buf); b_out = _DataBlob()
        ok = ctypes.windll.crypt32.CryptProtectData(
            ctypes.byref(b_in), None, None, None, None, 0, ctypes.byref(b_out))
        if ok:
            enc = bytes(ctypes.string_at(b_out.pbData, b_out.cbData))
            ctypes.windll.kernel32.LocalFree(b_out.pbData)
            return enc
    except Exception as e:
        logging.warning("DPAPI encrypt falhou: %s", e)
    return None

def _dpapi_decrypt(enc: bytes):
    try:
        buf   = ctypes.create_string_buffer(enc)
        b_in  = _DataBlob(len(enc), buf); b_out = _DataBlob()
        ok = ctypes.windll.crypt32.CryptUnprotectData(
            ctypes.byref(b_in), None, None, None, None, 0, ctypes.byref(b_out))
        if ok:
            dados = bytes(ctypes.string_at(b_out.pbData, b_out.cbData))
            ctypes.windll.kernel32.LocalFree(b_out.pbData)
            return dados
    except Exception as e:
        logging.warning("DPAPI decrypt falhou: %s", e)
    return None

# ── MSAL — copiado do Torre Web ───────────────────────────────────────────────
_TOKEN_CACHE_PATH = os.path.join(BASE_DIR, '.sp_token_cache.bin')

def _get_msal_app():
    import msal
    cache = msal.SerializableTokenCache()
    if os.path.exists(_TOKEN_CACHE_PATH):
        try:
            with open(_TOKEN_CACHE_PATH, 'rb') as f: raw = f.read()
            dec = _dpapi_decrypt(raw)
            if dec:
                cache.deserialize(dec.decode('utf-8'))
                logging.info("Token cache carregado (DPAPI)")
            else:
                cache.deserialize(raw.decode('utf-8'))
                logging.info("Token cache carregado (texto plano)")
        except Exception as e:
            logging.warning("Erro ao ler token cache: %s", e)
    return msal.PublicClientApplication(_AAD_CLIENT_ID, authority=_AAD_AUTHORITY, token_cache=cache), cache

def _salvar_cache(cache):
    try:
        if cache and cache.has_state_changed:
            dados = cache.serialize().encode('utf-8')
            enc   = _dpapi_encrypt(dados)
            with open(_TOKEN_CACHE_PATH, 'wb') as f:
                f.write(enc if enc else dados)
            logging.info("Token cache salvo")
    except Exception as e:
        logging.warning("Não foi possível salvar token cache: %s", e)

# ── AUTH STATUS — igual ao Torre Web ─────────────────────────────────────────
_sharepoint_status = {'ok': None, 'mensagem': '', 'device_code': None}

def _adquirir_token():
    """
    Tenta token silencioso do cache.
    Se não houver, inicia Device Flow, abre /auth e aguarda (bloqueante).
    Igual ao Torre Web.
    """
    pca, cache = _get_msal_app()

    # Tenta silencioso
    contas = pca.get_accounts()
    if contas:
        result = pca.acquire_token_silent(_SCOPES, account=contas[0])
        if result and 'access_token' in result:
            _salvar_cache(cache)
            return result['access_token']

    # Device Flow
    flow = pca.initiate_device_flow(scopes=_SCOPES)
    if 'user_code' not in flow:
        logging.error("Falha ao iniciar Device Flow: %s", flow)
        _sharepoint_status['ok'] = False
        _sharepoint_status['mensagem'] = flow.get('error_description', 'Erro no Device Flow')
        return None

    _sharepoint_status['device_code'] = {
        'user_code':  flow['user_code'],
        'verify_url': flow.get('verification_uri', 'https://microsoft.com/devicelogin'),
    }
    _sharepoint_status['ok'] = None
    _sharepoint_status['mensagem'] = 'Aguardando autenticação...'
    logging.info("Device Flow. Código: %s", flow['user_code'])

    # Abre /auth igual ao Torre Web
    try:
        webbrowser.open(f'http://127.0.0.1:{PORT}/auth')
    except Exception: pass

    # Aguarda (bloqueante)
    result = pca.acquire_token_by_device_flow(flow)
    _sharepoint_status['device_code'] = None

    if 'access_token' in result:
        _salvar_cache(cache)
        _sharepoint_status['ok'] = True
        _sharepoint_status['mensagem'] = '✓ Autenticado!'
        return result['access_token']

    logging.error("Device Flow falhou: %s", result.get('error_description'))
    _sharepoint_status['ok'] = False
    _sharepoint_status['mensagem'] = result.get('error_description', 'Falha na autenticação')
    return None

def _get_token():
    """Token silencioso apenas (para middleware de API). Não bloqueia."""
    try:
        pca, cache = _get_msal_app()
        contas = pca.get_accounts()
        if contas:
            result = pca.acquire_token_silent(_SCOPES, account=contas[0])
            if result and 'access_token' in result:
                _salvar_cache(cache)
                return result['access_token'], None
        return None, 'not_authenticated'
    except Exception as e:
        return None, str(e)

# ── USER CACHE ────────────────────────────────────────────────────────────────
_user_cache = {'user': None, 'ts': 0.0}
USER_CACHE_TTL = 300

def _decode_jwt_claims(token):
    """Decodifica o payload JWT sem validar assinatura (claims públicos)."""
    import base64 as _b64, json as _j
    try:
        payload = token.split('.')[1]
        payload += '=' * (4 - len(payload) % 4)
        return _j.loads(_b64.urlsafe_b64decode(payload))
    except Exception:
        return {}

def _get_current_user(token):
    """Extrai identidade do usuário direto do token JWT — sem chamar Graph /me."""
    now_t = time.time()
    if _user_cache['user'] and now_t - _user_cache['ts'] < USER_CACHE_TTL:
        return _user_cache['user']
    if not token:
        return None
    try:
        claims = _decode_jwt_claims(token)
        uid    = claims.get('oid') or claims.get('sub', '')
        name   = claims.get('name') or claims.get('given_name', '')
        email  = claims.get('upn') or claims.get('preferred_username', '') or claims.get('email', '')
        if not uid:
            return None
        user = {
            'id':       uid,
            'name':     name,
            'email':    email,
            'initials': ''.join(w[0].upper() for w in name.split()[:2]) or '?',
        }
        _user_cache['user'] = user
        _user_cache['ts']   = now_t
        return user
    except Exception:
        return None

def _clear_user_cache():
    _user_cache['user'] = None
    _user_cache['ts']   = 0.0

# ── SHAREPOINT HELPERS ────────────────────────────────────────────────────────
def _sp_file_url(rel_path):
    return f"{GRAPH}/sites/{SP_HOST}:/sites/{SP_SITE}:/drive/root:/{SP_FOLDER}/{rel_path}:/content"

def _sp_folder_url(folder_rel):
    return f"{GRAPH}/sites/{SP_HOST}:/sites/{SP_SITE}:/drive/root:/{SP_FOLDER}/{folder_rel}:/children"

def sp_get_json(rel_path):
    import requests as rq
    token, err = _get_token()
    if not token: return None, err
    try:
        r = rq.get(_sp_file_url(rel_path), headers={'Authorization': f'Bearer {token}'}, timeout=20)
        if r.status_code == 200: return r.json(), None
        if r.status_code == 404: return None, 'not_found'
        return None, f'HTTP {r.status_code}'
    except Exception as e: return None, str(e)

def sp_put_json(rel_path, data):
    import requests as rq
    token, err = _get_token()
    if not token: return False, err
    try:
        body = json.dumps(data, ensure_ascii=False, indent=2).encode('utf-8')
        r = rq.put(_sp_file_url(rel_path),
                   headers={'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'},
                   data=body, timeout=30)
        return r.status_code in (200, 201), f'HTTP {r.status_code}'
    except Exception as e: return False, str(e)

def sp_list_files(folder_rel):
    import requests as rq
    token, err = _get_token()
    if not token: return [], err
    try:
        r = rq.get(_sp_folder_url(folder_rel), headers={'Authorization': f'Bearer {token}'}, timeout=15)
        if r.status_code == 200: return r.json().get('value', []), None
        if r.status_code == 404: return [], None
        return [], f'HTTP {r.status_code}'
    except Exception as e: return [], str(e)

# ── DATA CACHE ────────────────────────────────────────────────────────────────
_data_cache = None
_cache_ts   = 0.0
CACHE_TTL   = 15
_data_lock  = threading.Lock()

def _default_data():
    return {"projects":[],"items":[],"flowcharts":[],"versions":[],"mapa":[],"users":[],"contacts":[]}

def load_data():
    global _data_cache, _cache_ts
    now_t = time.time()
    with _data_lock:
        if _data_cache is not None and now_t - _cache_ts < CACHE_TTL:
            return json.loads(json.dumps(_data_cache))
        data, err = sp_get_json('tracker_data.json')
        if data is None:
            if err == 'not_found':
                data = _default_data(); sp_put_json('tracker_data.json', data)
            elif os.path.exists(DATA_FILE):
                with open(DATA_FILE, 'r', encoding='utf-8') as f: data = json.load(f)
            else:
                data = _default_data()
        for k in _default_data(): data.setdefault(k, [])
        _data_cache = data; _cache_ts = now_t
    return json.loads(json.dumps(data))

def save_data(data):
    global _data_cache, _cache_ts
    with _data_lock: _data_cache = json.loads(json.dumps(data)); _cache_ts = time.time()
    ok, _ = sp_put_json('tracker_data.json', data)
    if not ok:
        with open(DATA_FILE, 'w', encoding='utf-8') as f: json.dump(data, f, ensure_ascii=False, indent=2)

# ── AUTH MIDDLEWARE ───────────────────────────────────────────────────────────
_PUBLIC_API = {'/api/auth/status', '/api/auth/logout'}

@app.before_request
def require_auth():
    if request.path.startswith('/api/') and request.path not in _PUBLIC_API:
        token, _ = _get_token()
        if not token:
            return jsonify({'error': 'Não autenticado', 'code': 'UNAUTHENTICATED'}), 401

_ROLES = ('admin', 'lider', 'membro')

def _get_user_role(user_id, data=None):
    """Retorna o papel do usuário no banco ('admin'|'lider'|'membro')."""
    if data is None: data = load_data()
    db_u = next((u for u in data['users'] if u['id'] == user_id), None)
    return db_u.get('role', 'membro') if db_u else 'membro'

def _admin_required():
    token, _ = _get_token()
    user = _get_current_user(token) if token else None
    if not user: return None, (jsonify({'error': 'Não autenticado'}), 401)
    data = load_data()
    if _get_user_role(user['id'], data) != 'admin':
        return None, (jsonify({'error': 'Apenas admins podem realizar esta ação'}), 403)
    return data, None

def _lider_or_admin_required():
    """Exige papel lider ou admin."""
    token, _ = _get_token()
    user = _get_current_user(token) if token else None
    if not user: return None, (jsonify({'error': 'Não autenticado'}), 401)
    data = load_data()
    if _get_user_role(user['id'], data) not in ('admin', 'lider'):
        return None, (jsonify({'error': 'Apenas líderes e admins podem realizar esta ação'}), 403)
    return data, None

# ── ROTAS PRINCIPAIS ──────────────────────────────────────────────────────────
@app.route('/')
def index():
    token, _ = _get_token()
    if not token:
        # Inicia auth em background e redireciona para /auth
        threading.Thread(target=_adquirir_token, daemon=True).start()
        time.sleep(0.8)
        return redirect('/auth')
    return render_template('index.html')

# ── AUTH — igual ao Torre Web ─────────────────────────────────────────────────
@app.route('/auth')
@limiter.limit("10 per minute")
def auth_page():
    """Mostra o device code — igual à /auth do Torre Web."""
    dc = _sharepoint_status.get('device_code')
    if not dc:
        # Sem device code pendente — inicia o fluxo e aguarda um pouco
        threading.Thread(target=_adquirir_token, daemon=True).start()
        time.sleep(1.2)
        dc = _sharepoint_status.get('device_code')
    if dc:
        return render_template('auth.html', device_code={
            'user_code':        dc['user_code'],
            'verification_uri': dc.get('verify_url', 'https://microsoft.com/devicelogin'),
        })
    # Se não tem device_code ainda (ex: erro), mostra placeholder
    return render_template('auth.html', device_code={
        'user_code': '------', 'verification_uri': 'https://microsoft.com/devicelogin'
    })

@app.route('/sharepoint-status')
def sharepoint_status():
    """Polling da /auth — igual ao Torre Web."""
    return jsonify({
        'ok':       _sharepoint_status.get('ok'),
        'mensagem': _sharepoint_status.get('mensagem', ''),
    })

# ── AUTH API ──────────────────────────────────────────────────────────────────
@app.route('/api/auth/status')
def auth_status():
    token, _ = _get_token()
    if token:
        user = _get_current_user(token)
        if user:
            try:
                data = load_data()
                db_u = next((u for u in data['users'] if u['id'] == user['id']), None)
                user['role'] = db_u.get('role', 'membro') if db_u else 'membro'
            except Exception: pass
        return jsonify({'connected': True, 'user': user})
    return jsonify({'connected': False})

@app.route('/api/auth/logout', methods=['POST'])
def auth_logout():
    _clear_user_cache()
    if os.path.exists(_TOKEN_CACHE_PATH): os.remove(_TOKEN_CACHE_PATH)
    _sharepoint_status['ok'] = None
    _sharepoint_status['device_code'] = None
    return jsonify({'ok': True})

# ── USERS ─────────────────────────────────────────────────────────────────────
@app.route('/api/me')
def get_me():
    token, _ = _get_token()
    user = _get_current_user(token) if token else None
    if not user: return jsonify({'error': 'Não foi possível obter usuário'}), 500
    data = load_data()
    db_u = next((u for u in data['users'] if u['id'] == user['id']), None)
    if not db_u:
        user['role'] = 'admin' if not data['users'] else 'membro'
        data['users'].append({**user}); save_data(data)
    else:
        user['role'] = db_u.get('role', 'membro')
    return jsonify(user)

@app.route('/api/users')
def get_users(): return jsonify(load_data().get('users', []))

@app.route('/api/users/<uid>/role', methods=['PUT'])
def set_user_role(uid):
    token, _ = _get_token()
    caller = _get_current_user(token) if token else None
    if not caller: return jsonify({'error': 'Não autenticado'}), 401
    data = load_data()
    caller_db = next((u for u in data['users'] if u['id'] == caller['id']), None)
    if not caller_db or caller_db.get('role') != 'admin':
        return jsonify({'error': 'Apenas admins podem alterar papéis'}), 403
    new_role = request.json.get('role')
    if new_role not in _ROLES: return jsonify({'error': 'Papel inválido'}), 400
    for u in data['users']:
        if u['id'] == uid: u['role'] = new_role; save_data(data); return jsonify(u)
    return jsonify({'error': 'Usuário não encontrado'}), 404

# ── CONTACTS ─────────────────────────────────────────────────────────────────
@app.route('/api/contacts', methods=['GET'])
def get_contacts():
    return jsonify(load_data().get('contacts', []))

@app.route('/api/contacts', methods=['POST'])
def add_contact():
    token, _ = _get_token()
    if not token: return jsonify({'error': 'Não autenticado'}), 401
    data = load_data()
    c = request.json or {}
    if not c.get('nome'): return jsonify({'error': 'Nome obrigatório'}), 400
    c['id']         = str(uuid.uuid4())
    c['created_at'] = now()
    data.setdefault('contacts', []).append(c)
    save_data(data)
    return jsonify(c), 201

@app.route('/api/contacts/<cid>', methods=['PUT'])
def update_contact(cid):
    token, _ = _get_token()
    if not token: return jsonify({'error': 'Não autenticado'}), 401
    data = load_data()
    for i, c in enumerate(data.get('contacts', [])):
        if c['id'] == cid:
            data['contacts'][i].update(request.json or {})
            save_data(data)
            return jsonify(data['contacts'][i])
    return jsonify({'error': 'not found'}), 404

@app.route('/api/contacts/<cid>', methods=['DELETE'])
def delete_contact(cid):
    data, err = _admin_required()
    if err: return err
    data['contacts'] = [c for c in data.get('contacts', []) if c['id'] != cid]
    save_data(data)
    return jsonify({'ok': True})

# ── PROJECTS ──────────────────────────────────────────────────────────────────
@app.route('/api/projects', methods=['GET'])
def get_projects(): return jsonify(load_data()['projects'])

@app.route('/api/projects', methods=['POST'])
def add_project():
    data, err = _admin_required()
    if err: return err
    p = request.json; p['id'] = str(uuid.uuid4()); p['created_at'] = now()
    data['projects'].append(p); save_data(data); return jsonify(p), 201

@app.route('/api/projects/<pid>', methods=['PUT'])
def update_project(pid):
    data, err = _admin_required()
    if err: return err
    for i, p in enumerate(data['projects']):
        if p['id'] == pid: data['projects'][i].update(request.json); save_data(data); return jsonify(data['projects'][i])
    return jsonify({'error': 'not found'}), 404

@app.route('/api/projects/<pid>', methods=['DELETE'])
def delete_project(pid):
    data, err = _admin_required()
    if err: return err
    data['projects']   = [p for p in data['projects']   if p['id'] != pid]
    data['items']      = [i for i in data['items']      if i.get('project_id') != pid]
    data['flowcharts'] = [f for f in data['flowcharts'] if f.get('project_id') != pid]
    data['versions']   = [v for v in data['versions']   if v.get('project_id') != pid]
    data['mapa']       = [m for m in data['mapa']       if m.get('project_id') != pid]
    save_data(data); return jsonify({'ok': True})

# ── ITEMS ─────────────────────────────────────────────────────────────────────
@app.route('/api/items', methods=['GET'])
def get_items():
    data = load_data(); pid = request.args.get('project')
    return jsonify([i for i in data['items'] if i.get('project_id') == pid] if pid else data['items'])

@app.route('/api/items', methods=['POST'])
def add_item():
    token, _ = _get_token()
    user  = _get_current_user(token) if token else None
    data  = load_data()
    item  = request.json
    item['id']         = str(uuid.uuid4())
    item['created_at'] = now()
    # Registra proprietário do item
    if user:
        item['owner_id']   = user['id']
        item['owner_name'] = user.get('name', '')
    if item.get('assignee_id') and not item.get('assignee_name'):
        u = next((u for u in data['users'] if u['id'] == item['assignee_id']), None)
        if u: item['assignee_name'] = u.get('name','')
    data['items'].append(item); save_data(data); return jsonify(item), 201

@app.route('/api/items/<iid>', methods=['PUT'])
def update_item(iid):
    token, _ = _get_token()
    user  = _get_current_user(token) if token else None
    if not user: return jsonify({'error': 'Não autenticado'}), 401
    data  = load_data()
    role  = _get_user_role(user['id'], data)
    is_privileged = role in ('admin', 'lider')  # admin e lider editam qualquer item
    payload = request.json
    if payload.get('assignee_id') and not payload.get('assignee_name'):
        u = next((u for u in data['users'] if u['id'] == payload['assignee_id']), None)
        if u: payload['assignee_name'] = u.get('name','')
    for i, item in enumerate(data['items']):
        if item['id'] == iid:
            # Só dono, lider ou admin pode editar
            if item.get('owner_id') and item['owner_id'] != user['id'] and not is_privileged:
                return jsonify({'error': 'Você não tem permissão para editar este item'}), 403
            data['items'][i].update(payload); save_data(data); return jsonify(data['items'][i])
    return jsonify({'error': 'not found'}), 404

@app.route('/api/items/<iid>', methods=['DELETE'])
def delete_item(iid):
    token, _ = _get_token()
    user  = _get_current_user(token) if token else None
    if not user: return jsonify({'error': 'Não autenticado'}), 401
    data  = load_data()
    is_privileged = _get_user_role(user['id'], data) in ('admin', 'lider')
    item  = next((i for i in data['items'] if i['id'] == iid), None)
    if not item: return jsonify({'error': 'not found'}), 404
    if item.get('owner_id') and item['owner_id'] != user['id'] and not is_privileged:
        return jsonify({'error': 'Você não tem permissão para apagar este item'}), 403
    data['items'] = [i for i in data['items'] if i['id'] != iid]
    save_data(data); return jsonify({'ok': True})

# ── FLOWCHARTS ────────────────────────────────────────────────────────────────
@app.route('/api/flowcharts', methods=['GET'])
def get_flowcharts():
    data = load_data(); pid = request.args.get('project')
    return jsonify([f for f in data['flowcharts'] if f.get('project_id') == pid] if pid else data['flowcharts'])

@app.route('/api/flowcharts', methods=['POST'])
def add_flowchart():
    data = load_data(); fc = request.json; fc['id'] = str(uuid.uuid4())
    data['flowcharts'].append(fc); save_data(data); return jsonify(fc), 201

@app.route('/api/flowcharts/<fid>', methods=['PUT'])
def update_flowchart(fid):
    data = load_data()
    for i, fc in enumerate(data['flowcharts']):
        if fc['id'] == fid:
            payload = request.json; payload['id'] = fid
            data['flowcharts'][i] = payload; save_data(data); return jsonify(data['flowcharts'][i])
    return jsonify({'error': 'not found'}), 404

@app.route('/api/flowcharts/<fid>', methods=['DELETE'])
def delete_flowchart(fid):
    data = load_data(); data['flowcharts'] = [f for f in data['flowcharts'] if f['id'] != fid]
    save_data(data); return jsonify({'ok': True})

# ── VERSIONS ──────────────────────────────────────────────────────────────────
@app.route('/api/versions', methods=['GET'])
def get_versions():
    data = load_data(); pid = request.args.get('project')
    return jsonify([v for v in data['versions'] if v.get('project_id') == pid] if pid else data['versions'])

@app.route('/api/versions', methods=['POST'])
def add_version():
    data = load_data(); v = request.json; v['id'] = str(uuid.uuid4()); v['created_at'] = now()
    data['versions'].append(v); save_data(data); return jsonify(v), 201

@app.route('/api/versions/<vid>', methods=['PUT'])
def update_version(vid):
    data = load_data()
    for i, v in enumerate(data['versions']):
        if v['id'] == vid: data['versions'][i].update(request.json); save_data(data); return jsonify(data['versions'][i])
    return jsonify({'error': 'not found'}), 404

@app.route('/api/versions/<vid>', methods=['DELETE'])
def delete_version(vid):
    data = load_data(); data['versions'] = [v for v in data['versions'] if v['id'] != vid]
    save_data(data); return jsonify({'ok': True})

# ── MAPA ──────────────────────────────────────────────────────────────────────
@app.route('/api/mapa', methods=['GET'])
def get_mapa():
    data = load_data(); pid = request.args.get('project')
    return jsonify([m for m in data['mapa'] if m.get('project_id') == pid] if pid else data['mapa'])

@app.route('/api/mapa', methods=['POST'])
def add_mapa():
    data = load_data(); m = request.json; m['id'] = str(uuid.uuid4()); m['created_at'] = now()
    data['mapa'].append(m); save_data(data); return jsonify(m), 201

@app.route('/api/mapa/<mid>', methods=['PUT'])
def update_mapa(mid):
    data = load_data()
    for i, m in enumerate(data['mapa']):
        if m['id'] == mid: data['mapa'][i].update(request.json); save_data(data); return jsonify(data['mapa'][i])
    return jsonify({'error': 'not found'}), 404

@app.route('/api/mapa/<mid>', methods=['DELETE'])
def delete_mapa(mid):
    data = load_data(); data['mapa'] = [m for m in data['mapa'] if m['id'] != mid]
    save_data(data); return jsonify({'ok': True})

# ── GIT ───────────────────────────────────────────────────────────────────────
@app.route('/api/git-log', methods=['GET'])
def git_log():
    path = request.args.get('path','').strip(); limit = int(request.args.get('limit',30))
    if not path or not os.path.isdir(path): return jsonify({'error': 'Caminho inválido'}), 400
    try:
        tag_r = subprocess.run(['git','tag','--sort=-version:refname','--format=%(refname:short)|%(creatordate:short)|%(subject)'],cwd=path,capture_output=True,text=True,timeout=10)
        tags = [{'ref':p[0],'date':p[1] if len(p)>1 else '','subject':p[2] if len(p)>2 else ''} for p in [l.split('|') for l in tag_r.stdout.strip().splitlines()] if p[0]]
        log_r = subprocess.run(['git','log','--pretty=format:%h|%ad|%s|%D','--date=short',f'-{limit}'],cwd=path,capture_output=True,text=True,timeout=10)
        commits = [{'hash':p[0] if len(p)>0 else '','date':p[1] if len(p)>1 else '','subject':p[2] if len(p)>2 else '','refs':p[3] if len(p)>3 else ''} for p in [l.split('|') for l in log_r.stdout.strip().splitlines()]]
        return jsonify({'tags': tags[:limit], 'commits': commits})
    except Exception as e: return jsonify({'error': str(e)}), 500

# ── SEED DEMO DATA ────────────────────────────────────────────────────────────
@app.route('/api/seed', methods=['POST'])
def seed():
    data_req, err = _admin_required()
    if err: return err
    data = data_req
    if data['projects']: return jsonify({'ok': True, 'count': 0})
    today = now()

    # ── Projeto 1: Torre Web ──────────────────────────────────────────────────
    pid1 = str(uuid.uuid4())
    data['projects'].append({'id':pid1,'name':'Torre Web','color':'#22d3ee','icon':'🗼','created_at':today})

    # ── Projeto 2: Tracker ────────────────────────────────────────────────────
    pid2 = str(uuid.uuid4())
    data['projects'].append({'id':pid2,'name':'Tracker','color':'#a78bfa','icon':'📋','created_at':today})

    # ── Contatos demo ─────────────────────────────────────────────────────────
    CONTACTS = [
        {"nome":"Fernanda Lima",      "empresa":"Arezzo","cargo":"Gerente de Compras — Calçados",   "email":"fernanda.lima@arezzo.com.br",    "telefone":"11 99101-2030"},
        {"nome":"Carlos Eduardo Mota","empresa":"Arezzo","cargo":"Coordenador de Planejamento",      "email":"carlos.mota@arezzo.com.br",      "telefone":"11 99102-4050"},
        {"nome":"Priya Nair",         "empresa":"Arezzo","cargo":"Head de Dados e Analytics",        "email":"priya.nair@arezzo.com.br",       "telefone":"11 99103-6070"},
        {"nome":"Marcos Teixeira",    "empresa":"Arezzo","cargo":"Gerente de TI",                    "email":"marcos.teixeira@arezzo.com.br",  "telefone":"11 99104-8090"},
        {"nome":"Beatriz Andrade",    "empresa":"Arezzo","cargo":"Analista de Processos",            "email":"beatriz.andrade@arezzo.com.br",  "telefone":"11 99105-1020"},
        {"nome":"Rafael Souza",       "empresa":"Parceiro — Agência Digital","cargo":"Gerente de Projetos","email":"rafael.souza@parceiro.com","telefone":"11 98201-3040"},
    ]
    for c in CONTACTS:
        c['id'] = str(uuid.uuid4()); c['created_at'] = today
        data.setdefault('contacts', []).append(c)

    # Indexar contatos por nome para referenciar nos itens
    cmap = {c['nome']: c for c in data['contacts']}

    ITEMS = [
        # Torre Web — concluídos
        {"project_id":pid1,"titulo":"PME/Cobertura/Giro — fix do toggle de métricas","descricao":"Toggle alternava entre PME e Cobertura sem persistir o estado selecionado ao trocar de loja.","tipo":"bug","status":"concluido","categoria":"UX / Interface","prioridade":"P1","sprint":"Jul W1","esforco":"S","solicitante_nome":"Fernanda Lima","solicitante_id":cmap["Fernanda Lima"]["id"]},
        {"project_id":pid1,"titulo":"Geração de síntese na análise executiva via LLM","descricao":"Endpoint que gera texto de síntese automático usando GPT-4o com os dados da visão executiva.","tipo":"feature","status":"concluido","categoria":"Dados","prioridade":"P2","sprint":"Jul W1","esforco":"L","solicitante_nome":"Priya Nair","solicitante_id":cmap["Priya Nair"]["id"]},
        {"project_id":pid1,"titulo":"Cache de resultados SharePoint — TTL 10 min","descricao":"Resultados de consulta ao SharePoint agora são cacheados por 10 minutos para reduzir latência.","tipo":"melhoria","status":"concluido","categoria":"Performance","prioridade":"P2","sprint":"Jul W2","esforco":"M"},
        {"project_id":pid1,"titulo":"Exportação Excel da visão de cobertura","descricao":"Botão exportar na tela de cobertura gera .xlsx com todas as lojas e métricas filtradas.","tipo":"feature","status":"concluido","categoria":"Dados","prioridade":"P3","sprint":"Jun W4","esforco":"M","solicitante_nome":"Carlos Eduardo Mota","solicitante_id":cmap["Carlos Eduardo Mota"]["id"]},

        # Torre Web — em andamento
        {"project_id":pid1,"titulo":"Filtro de cluster por regional no mapa de calor","descricao":"Adicionar dropdown de regional para filtrar o mapa de calor de cobertura por agrupamento geográfico.","tipo":"feature","status":"andamento","categoria":"UX / Interface","prioridade":"P1","sprint":"Jul W3","esforco":"M","solicitante_nome":"Fernanda Lima","solicitante_id":cmap["Fernanda Lima"]["id"],"fonte_nome":"Beatriz Andrade","fonte_id":cmap["Beatriz Andrade"]["id"]},
        {"project_id":pid1,"titulo":"Refactor da camada de acesso ao SharePoint","descricao":"Separar sp_get_json e sp_list_files em módulo próprio com retry automático e logging estruturado.","tipo":"refactor","status":"andamento","categoria":"Tecnologia","prioridade":"P2","sprint":"Jul W3","esforco":"L"},
        {"project_id":pid1,"titulo":"Dashboard de acompanhamento por comprador","descricao":"Nova visão que agrega métricas de PME, cobertura e giro por comprador responsável.","tipo":"feature","status":"andamento","categoria":"Dados","prioridade":"P1","sprint":"Jul W3","esforco":"XL","solicitante_nome":"Carlos Eduardo Mota","solicitante_id":cmap["Carlos Eduardo Mota"]["id"],"fonte_nome":"Priya Nair","fonte_id":cmap["Priya Nair"]["id"]},

        # Torre Web — backlog
        {"project_id":pid1,"titulo":"Alertas automáticos de ruptura via email","descricao":"Envio diário de email com lojas em ruptura crítica (cobertura < 5 dias) para o time de compras.","tipo":"feature","status":"backlog","categoria":"Integração","prioridade":"P1","sprint":"Jul W4","esforco":"L","solicitante_nome":"Fernanda Lima","solicitante_id":cmap["Fernanda Lima"]["id"]},
        {"project_id":pid1,"titulo":"Modo comparativo — duas safras lado a lado","descricao":"Permitir comparar métricas de duas safras distintas na mesma visualização.","tipo":"feature","status":"backlog","categoria":"UX / Interface","prioridade":"P2","sprint":"Ago W1","esforco":"XL"},
        {"project_id":pid1,"titulo":"Performance da query de giro — índice na tabela de movimentação","descricao":"Query de giro ultrapassa 4s em lojas com alto volume. Adicionar índice composto.","tipo":"melhoria","status":"backlog","categoria":"Performance","prioridade":"P2","sprint":"Jul W4","esforco":"S","fonte_nome":"Marcos Teixeira","fonte_id":cmap["Marcos Teixeira"]["id"]},
        {"project_id":pid1,"titulo":"Tratamento de erro quando SharePoint está offline","descricao":"Exibir mensagem amigável e dados do último cache válido quando SharePoint retorna 503.","tipo":"bug","status":"backlog","categoria":"UX / Interface","prioridade":"P3","sprint":"Ago W1","esforco":"S"},

        # Tracker — concluídos
        {"project_id":pid2,"titulo":"Sistema de autenticação via Device Code (Torre Web pattern)","descricao":"Substituir modal de login pelo fluxo de device code da Microsoft, com tela /auth dedicada e cache DPAPI.","tipo":"feature","status":"concluido","categoria":"Tecnologia","prioridade":"P1","sprint":"Jul W2","esforco":"L","solicitante_nome":"Marcos Teixeira","solicitante_id":cmap["Marcos Teixeira"]["id"]},
        {"project_id":pid2,"titulo":"Controle de propriedade de itens (owner_id)","descricao":"Cada item registra quem criou. Somente o dono ou admin pode editar/apagar. Frontend esconde botões.","tipo":"feature","status":"concluido","categoria":"Processo","prioridade":"P1","sprint":"Jul W3","esforco":"M"},
        {"project_id":pid2,"titulo":"Remoção da aba Agenda Outlook","descricao":"Agenda removida por depender de Calendars.Read bloqueado no tenant. Substituída por Dashboard.","tipo":"melhoria","status":"concluido","categoria":"UX / Interface","prioridade":"P2","sprint":"Jul W3","esforco":"S"},

        # Tracker — em andamento
        {"project_id":pid2,"titulo":"Dashboard de resumo por especialista","descricao":"Nova aba com stats gerais, cards por especialista com suas demandas principais e itens críticos.","tipo":"feature","status":"andamento","categoria":"UX / Interface","prioridade":"P1","sprint":"Jul W3","esforco":"M"},

        # Tracker — backlog
        {"project_id":pid2,"titulo":"Notificações de prazo próximo (toast + badge)","descricao":"Mostrar toast na abertura quando há itens com prazo em 3 dias ou menos no projeto ativo.","tipo":"feature","status":"backlog","categoria":"UX / Interface","prioridade":"P2","sprint":"Jul W4","esforco":"S"},
        {"project_id":pid2,"titulo":"Export PDF do roadmap/mapa","descricao":"Botão para exportar o mapa de projetos como PDF com layout landscape.","tipo":"feature","status":"backlog","categoria":"Dados","prioridade":"P3","sprint":"Ago W1","esforco":"M"},
        {"project_id":pid2,"titulo":"Filtro de assignee no Kanban e Processos","descricao":"Dropdown para filtrar cards por responsável em todas as visões de board.","tipo":"melhoria","status":"backlog","categoria":"UX / Interface","prioridade":"P2","sprint":"Jul W4","esforco":"S"},
    ]
    total = 0
    for item in ITEMS:
        item['id'] = str(uuid.uuid4()); item['created_at'] = today
        data['items'].append(item); total += 1
    save_data(data); return jsonify({'ok': True, 'count': total, 'contacts': len(CONTACTS)})

# ── SEEDS ─────────────────────────────────────────────────────────────────────
import shutil

def safe_seed_filename(name):
    safe = name.replace('/','').replace('\\','').replace('..','').strip()
    if not safe.endswith('.json'): safe += '.json'
    return safe

def get_seed_data(name):
    filename = safe_seed_filename(name)
    data, err = sp_get_json(f'{SP_SEEDS}/{filename}')
    if data: return data, None
    fp = os.path.join(SEED_DIR, filename)
    if os.path.exists(fp):
        with open(fp, encoding='utf-8') as f: return json.load(f), None
    return None, f'Seed {name} não encontrado'

def _seed_meta(filename, d):
    meta = d.get('_meta',{}); items = d.get('items',[])
    especialista = meta.get('especialista') or _parse_especialista(filename) or ''
    return {'filename':filename,'name':filename.replace('.json',''),
            'created_at':meta.get('created_at',''),'especialista':especialista,
            'n_projects':len(d.get('projects',[])),'n_total':len(items),
            'n_finalizados':len([i for i in items if i.get('status') in STATUS_FINALIZADOS]),
            'n_pendentes':len([i for i in items if i.get('status') in STATUS_PENDENTES]),
            'n_outros':len([i for i in items if i.get('status') in STATUS_OUTROS]),'size_kb':0}

# Cache em memória para a lista de seeds do SharePoint
# {'data': [...] | None, 'ts': float}
_seeds_cache: dict = {'data': None, 'ts': 0.0}
_SEEDS_CACHE_TTL = 0  # 0 = só atualiza via botão manual (/api/seeds/refresh)

def _fetch_seeds_from_sp():
    """Busca lista de seeds do SharePoint e atualiza o cache. Retorna lista."""
    sp_files, _ = sp_list_files(SP_SEEDS)
    if not sp_files:
        return None
    seeds = []
    for f in sorted(sp_files, key=lambda x: x.get('name',''), reverse=True):
        fname = f.get('name','')
        if not fname.endswith('.json'): continue
        d, _ = sp_get_json(f'{SP_SEEDS}/{fname}')
        if d: seeds.append(_seed_meta(fname, d))
    _seeds_cache['data'] = seeds
    _seeds_cache['ts'] = time.time()
    return seeds

@app.route('/api/seeds', methods=['GET'])
def list_seeds():
    # Usa cache se disponível (atualização é sempre manual via /api/seeds/refresh)
    if _seeds_cache['data'] is not None:
        return jsonify(_seeds_cache['data'])
    # Primeira chamada: tenta SP; se falhar, cai para arquivos locais
    sp_seeds = _fetch_seeds_from_sp()
    if sp_seeds is not None:
        return jsonify(sp_seeds)
    # Fallback local
    seeds = []
    for fn in sorted(os.listdir(SEED_DIR), reverse=True):
        if not fn.endswith('.json'): continue
        fp = os.path.join(SEED_DIR, fn)
        try:
            with open(fp, encoding='utf-8') as f: d = json.load(f)
            meta = d.get('_meta',{}); items = d.get('items',[])
            seeds.append({'filename':fn,'name':fn.replace('.json',''),'created_at':meta.get('created_at',''),
                          'n_projects':len(d.get('projects',[])),'n_total':len(items),
                          'n_finalizados':len([i for i in items if i.get('status') in STATUS_FINALIZADOS]),
                          'n_pendentes':len([i for i in items if i.get('status') in STATUS_PENDENTES]),
                          'n_outros':len([i for i in items if i.get('status') in STATUS_OUTROS]),
                          'size_kb':round(os.path.getsize(fp)/1024,1)})
        except Exception: pass
    return jsonify(seeds)

@app.route('/api/seeds/refresh', methods=['POST'])
@limiter.limit("5 per minute")
def refresh_seeds():
    """Força re-busca da lista de seeds do SharePoint (ação manual do usuário)."""
    token = require_auth()
    if token is not None: return token
    _seeds_cache['data'] = None
    _seeds_cache['ts'] = 0.0
    sp_seeds = _fetch_seeds_from_sp()
    if sp_seeds is not None:
        return jsonify({'ok': True, 'source': 'sharepoint', 'count': len(sp_seeds)})
    return jsonify({'ok': False, 'source': 'none', 'msg': 'Não foi possível conectar ao SharePoint'}), 503

@app.route('/api/seeds', methods=['POST'])
@limiter.limit("10 per hour")
def criar_seed():
    data_req, err = _lider_or_admin_required()
    if err: return err
    body = request.json or {}
    especialista = (body.get('especialista') or '').strip() or None
    name = body.get('name', seed_name_auto(especialista)).strip().upper()
    snap = json.loads(json.dumps(load_data()))
    snap['_meta'] = {'name':name,'especialista':especialista or '','created_at':datetime.now().strftime('%d/%m/%Y %H:%M')}
    filename = safe_seed_filename(name)
    ok, _ = sp_put_json(f'{SP_SEEDS}/{filename}', snap)
    if not ok:
        with open(os.path.join(SEED_DIR, filename), 'w', encoding='utf-8') as f:
            json.dump(snap, f, ensure_ascii=False, indent=2)
    return jsonify({'ok':True,'name':name,'filename':filename,'especialista':especialista or ''})

@app.route('/api/seeds/<name>', methods=['GET'])
def get_seed(name):
    d, err = get_seed_data(name)
    if d is None: return jsonify({'error': err or 'not found'}), 404
    return jsonify(d)

@app.route('/api/seeds/<name>/restaurar', methods=['POST'])
def restaurar_seed(name):
    data_req, err = _lider_or_admin_required()
    if err: return err
    snap, serr = get_seed_data(name)
    if snap is None: return jsonify({'error': serr or 'not found'}), 404
    body = request.json or {}
    bak_name = 'PRE_RESTORE_' + datetime.now().strftime('%Y%m%dT%H%M%S')
    data_atual = load_data()
    bak_snap = json.loads(json.dumps(data_atual))
    bak_snap['_meta'] = {'name':'backup_auto','created_at':datetime.now().strftime('%d/%m/%Y %H:%M')}
    bak_filename = safe_seed_filename(bak_name)
    ok, _ = sp_put_json(f'{SP_SEEDS}/{bak_filename}', bak_snap)
    if not ok:
        with open(os.path.join(SEED_DIR, bak_filename), 'w', encoding='utf-8') as f:
            json.dump(bak_snap, f, ensure_ascii=False, indent=2)
    snap.pop('_meta', None)
    snap['users'] = data_atual.get('users',[])
    if body.get('levar_inacabados', False):
        snap_ids = {i['id'] for i in snap.get('items',[])}
        extras = [i for i in data_atual.get('items',[]) if i.get('status') in STATUS_PENDENTES and i['id'] not in snap_ids]
        snap.setdefault('items',[]).extend(extras)
    save_data(snap)
    return jsonify({'ok':True,'backup':bak_filename,'restored':True})

@app.route('/api/seeds/fechar', methods=['POST'])
def fechar_projeto():
    data_req, err = _lider_or_admin_required()
    if err: return err
    body = request.json or {}
    especialista = (body.get('especialista') or '').strip() or None
    name = body.get('name', seed_name_auto(especialista)).strip().upper()
    levar_inacab = body.get('levar_inacabados', True)
    data = data_req
    snap = json.loads(json.dumps(data))
    snap['_meta'] = {'name':name,'especialista':especialista or '','created_at':datetime.now().strftime('%d/%m/%Y %H:%M')}
    filename = safe_seed_filename(name)
    ok, _ = sp_put_json(f'{SP_SEEDS}/{filename}', snap)
    if not ok:
        with open(os.path.join(SEED_DIR, filename), 'w', encoding='utf-8') as f:
            json.dump(snap, f, ensure_ascii=False, indent=2)
    data["items"] = [i for i in data["items"] if i.get("status") in STATUS_PENDENTES] if levar_inacab else []
    save_data(data)
    return jsonify({"ok":True,"name":name,"n_levados":len(data["items"]),"especialista":especialista or ''})

@app.route('/api/seeds/sp/upload/<name>', methods=['POST'])
@limiter.limit("20 per hour")
def sp_upload_seed(name):
    data_req, err = _lider_or_admin_required()
    if err: return err
    filename = safe_seed_filename(name)
    fp = os.path.join(SEED_DIR, filename)
    if not os.path.exists(fp):
        return jsonify({'error': 'Seed local nao encontrado'}), 404
    try:
        with open(fp, encoding='utf-8') as f: d = json.load(f)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    ok, msg = sp_put_json(f'{SP_SEEDS}/{filename}', d)
    if ok:
        _seeds_cache['data'] = None
        return jsonify({'ok': True, 'filename': filename, 'msg': 'Enviado para SharePoint'})
    return jsonify({'ok': False, 'msg': msg or 'Falha no upload'}), 503

@app.route('/api/seeds/sp/download/<name>', methods=['POST'])
@limiter.limit("20 per hour")
def sp_download_seed(name):
    token, err = _get_token()
    if not token: return jsonify({'error': 'Nao autenticado'}), 401
    filename = safe_seed_filename(name)
    d, sp_err = sp_get_json(f'{SP_SEEDS}/{filename}')
    if d is None:
        return jsonify({'error': sp_err or 'Seed nao encontrado no SharePoint'}), 404
    fp = os.path.join(SEED_DIR, filename)
    try:
        with open(fp, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    return jsonify({'ok': True, 'filename': filename, 'msg': 'Baixado do SharePoint'})

@app.route("/api/shutdown", methods=["POST"])
def shutdown_server():
    import signal
    threading.Timer(0.3, lambda: os.kill(os.getpid(), signal.SIGTERM)).start()
    return jsonify({"ok": True, "msg": "Servidor encerrando..."})

if __name__ == '__main__':
    threading.Timer(1.2, lambda: webbrowser.open(f'http://localhost:{PORT}')).start()
    print(f'\n Tracker V3.9 iniciando em http://localhost:{PORT}\n')
    app.run(host=HOST, port=PORT, debug=False)
