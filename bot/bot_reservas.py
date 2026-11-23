# -*- coding: utf-8 -*-
"""
Bot de reservas — Pedidos de Material · Astillero TASA
=======================================================
Hace de puente entre la app web (Firebase) y la carpeta compartida de reservas:

  1. EXPORTAR   Toma las solicitudes en estado "Validada" de los sistemas habilitados
                (por defecto TUBERIAS), genera el Excel "Pedido" (mismo formato que la app)
                y lo deja en  <raíz>\\<PROYECTO>\\<n. SISTEMA>\\Solicitados.
                La solicitud pasa a "Exportada" (Cargada) en la app.
  2. EJECUTAR   Corre el Código_reserva_2.exe de esa carpeta (carga en SAP, mueve a
                Reservados y envía el PDF), uno a la vez.
  3. LEER       Revisa las carpetas Reservados; cada Excel nuevo se asocia a su solicitud
                (por proyecto + fecha de solicitud) y actualiza Reserva, Posición y Estado
                (Reservado / Bloqueado), igual que "Importar Excel devuelto" en la app.
  4. INFORMAR   Publica su estado en la app (Gestión → Bot) y atiende "Ejecutar ahora".

Uso:  python bot_reservas.py              (ciclo continuo)
      python bot_reservas.py --una-vez    (un solo ciclo y termina)
      python bot_reservas.py --prueba     (no escribe nada: solo muestra qué haría)
"""
import os, sys, re, json, time, glob, socket, getpass, argparse, traceback, subprocess, unicodedata, configparser, logging, math
from logging.handlers import RotatingFileHandler
from datetime import datetime, timedelta, timezone

try:
    import requests
    import openpyxl
    from openpyxl.utils import get_column_letter
except ImportError as e:  # pragma: no cover
    print('Falta instalar dependencias. Ejecute instalar.bat (pip install requests openpyxl).\n', e)
    input('Enter para salir…'); sys.exit(1)

VERSION = '1.0'
ESPERA_COPIA = float(os.environ.get('BOT_ESPERA_COPIA', 20))
BASE = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
log = logging.getLogger('bot')

# ───────────────────────── utilidades ─────────────────────────
def norm(s):
    s = unicodedata.normalize('NFD', str(s if s is not None else ''))
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn').lower()
    return re.sub(r'[^a-z0-9]', '', s)

def now_iso():
    return datetime.now(timezone.utc).isoformat(timespec='milliseconds').replace('+00:00', 'Z')

def parse_iso(s):
    try:
        return datetime.fromisoformat(str(s).replace('Z', '+00:00'))
    except Exception:
        return None

def split_mails(t):
    out = []
    for x in re.split(r'[;,\s]+', str(t or '')):
        x = x.strip().lower()
        if re.fullmatch(r'\S+@\S+\.\S+', x): out.append(x)
    return out

def uniq(l):
    seen, out = set(), []
    for x in l:
        if x not in seen: seen.add(x); out.append(x)
    return out

def cell_str(v):
    """Equivalente a cell() de la app: texto recortado (números enteros sin .0)."""
    if v is None: return ''
    if isinstance(v, datetime): return v
    if isinstance(v, float) and v.is_integer(): return str(int(v))
    return str(v).strip()

# ───────────────────────── configuración ─────────────────────────
DEF_CFG = {
    'firebase': {'api_key': '', 'project_id': 'pedidos-tasa', 'dominio': 'pedidos-tasa.app', 'usuario': 'bot', 'clave': ''},
    'carpetas': {'raices': r'Z:\Diseño\RESERVAS\Reservas\{año}; Z:\Diseño\RESERVAS\Reservas\{año_anterior}',
                 'sistemas': '*', 'carpeta_especialidad': 'TUBERIAS', 'subcarpeta_solicitados': 'Solicitados', 'subcarpeta_reservados': 'Reservados',
                 'ejecutable': 'C*digo_reserva*.exe'},
    'proceso': {'intervalo_minutos': '15', 'exportar': 'si', 'ejecutar_exe': 'no', 'leer_reservados': 'si',
                'tiempo_max_exe_minutos': '60', 'operador_sap': '', 'espera_respuesta_minutos': '10',
                'argumentos_exe': '--auto --operador {operador} --espera {espera}', 'reintentar_minutos': '60', 'procesar_desde': '', 'modo_prueba': 'no', 'zona_horaria': '-5'},
    'red': {'proxy': ''},
}

class Cfg:
    def automatiza(self, sistema):
        return self.todos or norm(sistema) in self.sis_norm

    def carpeta_de(self, sistema):
        """Nombre de la carpeta de especialidad (p. ej. '2. TUBERIAS') donde va una solicitud de ese sistema."""
        return self.mapa_carpeta.get(norm(sistema)) or self.carpeta_def or sistema

    def carpetas_esp(self):
        return list(dict.fromkeys([v for v in [self.carpeta_def, *self.mapa_carpeta.values()] if v] or
                                  ([] if self.todos else self.sistemas)))

    def __init__(self, path, prueba=False):
        cp = configparser.ConfigParser(interpolation=None, inline_comment_prefixes=(';', '#'))
        cp.optionxform = str
        cp.read_dict(DEF_CFG)
        if not os.path.exists(path):
            raise SystemExit(f'No existe {path}. Copie config.ini.ejemplo como config.ini y complételo.')
        for enc in ('utf-8-sig', 'cp1252'):
            try:
                with open(path, encoding=enc) as f: cp.read_file(f)
                break
            except UnicodeDecodeError:
                continue
        g = lambda s, k: cp.get(s, k, fallback=DEF_CFG.get(s, {}).get(k, '')).strip()
        si = lambda s, k: g(s, k).lower() in ('si', 'sí', 'true', '1', 'yes')
        self.api_key = g('firebase', 'api_key'); self.project = g('firebase', 'project_id')
        self.dominio = g('firebase', 'dominio'); self.usuario = g('firebase', 'usuario').lower(); self.clave = g('firebase', 'clave')
        hoy = datetime.now()
        self.raices = [r.strip().replace('{año}', str(hoy.year)).replace('{ano}', str(hoy.year))
                       .replace('{año_anterior}', str(hoy.year - 1)).replace('{ano_anterior}', str(hoy.year - 1))
                       for r in g('carpetas', 'raices').split(';') if r.strip()]
        self.sistemas = [x.strip() for x in re.split(r'[;,]', g('carpetas', 'sistemas')) if x.strip()] or ['*']
        self.todos = '*' in self.sistemas
        self.carpeta_def = g('carpetas', 'carpeta_especialidad')
        self.mapa_carpeta = {}
        if cp.has_section('carpeta_por_sistema'):
            for k, v in cp.items('carpeta_por_sistema'):
                if v.strip(): self.mapa_carpeta[norm(k)] = v.strip()
        self.sub_sol = g('carpetas', 'subcarpeta_solicitados'); self.sub_res = g('carpetas', 'subcarpeta_reservados')
        self.exe_patron = g('carpetas', 'ejecutable')
        self.intervalo = max(1, int(float(g('proceso', 'intervalo_minutos') or 15)))
        self.exportar = si('proceso', 'exportar'); self.ejecutar_exe = si('proceso', 'ejecutar_exe')
        self.leer_res = si('proceso', 'leer_reservados')
        self.exe_timeout = max(1, int(float(g('proceso', 'tiempo_max_exe_minutos') or 60))) * 60
        self.operador = g('proceso', 'operador_sap').upper()
        self.espera = max(1, int(float(g('proceso', 'espera_respuesta_minutos') or 10)))
        self.exe_args = g('proceso', 'argumentos_exe').replace('{operador}', self.operador or 'AIDROGO').replace('{espera}', str(self.espera)).split()
        self.reintentar = int(float(g('proceso', 'reintentar_minutos') or 0))
        self.procesar_desde = g('proceso', 'procesar_desde')
        self.prueba = prueba or si('proceso', 'modo_prueba')
        self.tz = timezone(timedelta(hours=float(g('proceso', 'zona_horaria') or -5)))
        self.proxy = g('red', 'proxy')
        self.alias = {}
        self.sis_norm = {norm(x) for x in self.sistemas}
        if cp.has_section('alias_proyectos'):
            for k, v in cp.items('alias_proyectos'):
                if v.strip(): self.alias[norm(k)] = v.strip()
        self.desc = ('todos los sistemas' if self.todos else ', '.join(self.sistemas)) + f' → carpeta {self.carpeta_def or "(la del sistema)"}'
        if not self.api_key or not self.clave:
            raise SystemExit('Complete api_key y clave en config.ini (sección [firebase]).')

# ───────────────────────── Firestore REST ─────────────────────────
def enc(v):
    if v is None: return {'nullValue': None}
    if isinstance(v, bool): return {'booleanValue': v}
    if isinstance(v, int): return {'integerValue': str(v)}
    if isinstance(v, float): return {'doubleValue': v}
    if isinstance(v, dict): return {'mapValue': {'fields': {k: enc(x) for k, x in v.items()}}}
    if isinstance(v, (list, tuple)): return {'arrayValue': {'values': [enc(x) for x in v]}}
    return {'stringValue': str(v)}

def dec(v):
    if 'nullValue' in v: return None
    if 'booleanValue' in v: return v['booleanValue']
    if 'integerValue' in v: return int(v['integerValue'])
    if 'doubleValue' in v: return float(v['doubleValue'])
    if 'stringValue' in v: return v['stringValue']
    if 'timestampValue' in v: return v['timestampValue']
    if 'mapValue' in v: return {k: dec(x) for k, x in (v['mapValue'].get('fields') or {}).items()}
    if 'arrayValue' in v: return [dec(x) for x in (v['arrayValue'].get('values') or [])]
    if 'referenceValue' in v: return v['referenceValue']
    if 'geoPointValue' in v: return v['geoPointValue']
    if 'bytesValue' in v: return v['bytesValue']
    return None

class Conflicto(Exception):
    pass

class FS:
    AUTH = 'https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key='
    REFRESH = 'https://securetoken.googleapis.com/v1/token?key='
    BASE_URL = 'https://firestore.googleapis.com/v1/projects/{}/databases/(default)/documents'

    def __init__(self, cfg):
        self.c = cfg; self.tok = None; self.ref = None; self.exp = 0
        self.ses = requests.Session()
        if cfg.proxy: self.ses.proxies = {'http': cfg.proxy, 'https': cfg.proxy}
        self.base = (os.environ.get('BOT_FS_BASE') or self.BASE_URL.format(cfg.project))
        self.auth_url = os.environ.get('BOT_AUTH_URL')

    def _login(self):
        email = f'{self.c.usuario}@{self.c.dominio}'
        url = (self.auth_url + '/login') if self.auth_url else self.AUTH + self.c.api_key
        r = self.ses.post(url, json={'email': email, 'password': self.c.clave, 'returnSecureToken': True}, timeout=30)
        if r.status_code != 200:
            msg = r.text
            try: msg = r.json()['error']['message']
            except Exception: pass
            raise RuntimeError(f'No se pudo iniciar sesión como "{self.c.usuario}": {msg}')
        j = r.json(); self.tok = j['idToken']; self.ref = j.get('refreshToken'); self.exp = time.time() + int(j.get('expiresIn', 3600))

    def _token(self):
        if not self.tok or time.time() > self.exp - 300:
            if self.ref and not self.auth_url:
                try:
                    r = self.ses.post(self.REFRESH + self.c.api_key, data={'grant_type': 'refresh_token', 'refresh_token': self.ref}, timeout=30)
                    if r.status_code == 200:
                        j = r.json(); self.tok = j['id_token']; self.ref = j['refresh_token']; self.exp = time.time() + int(j.get('expires_in', 3600)); return self.tok
                except requests.RequestException:
                    pass
            self._login()
        return self.tok

    def _req(self, method, url, **kw):
        for intento in range(3):
            try:
                r = self.ses.request(method, url, headers={'Authorization': 'Bearer ' + self._token()}, timeout=60, **kw)
            except requests.RequestException as e:
                if intento == 2: raise RuntimeError(f'Sin conexión con Firebase: {e}')
                time.sleep(5 * (intento + 1)); continue
            if r.status_code == 401 and intento < 2:
                self.tok = None; continue
            if r.status_code in (500, 502, 503, 504) and intento < 2:
                time.sleep(5 * (intento + 1)); continue
            return r
        return r

    def listar(self, col):
        out, tok = [], None
        while True:
            params = {'pageSize': 300}
            if tok: params['pageToken'] = tok
            r = self._req('GET', f'{self.base}/{col}', params=params)
            if r.status_code != 200: raise RuntimeError(f'Error leyendo {col}: {r.status_code} {r.text[:300]}')
            j = r.json()
            for d in j.get('documents', []):
                did = d['name'].rsplit('/', 1)[1]
                data = {k: dec(v) for k, v in (d.get('fields') or {}).items()}
                data['id'] = did; data['_ut'] = d.get('updateTime')
                out.append(data)
            tok = j.get('nextPageToken')
            if not tok: return out

    def leer(self, path):
        r = self._req('GET', f'{self.base}/{path}')
        if r.status_code == 404: return None
        if r.status_code != 200: raise RuntimeError(f'Error leyendo {path}: {r.status_code} {r.text[:300]}')
        d = r.json(); data = {k: dec(v) for k, v in (d.get('fields') or {}).items()}
        data['id'] = path.rsplit('/', 1)[1]; data['_ut'] = d.get('updateTime'); return data

    def escribir(self, path, data, campos=None, update_time=None):
        """campos=None reemplaza el documento; con campos solo actualiza esos. update_time = control de concurrencia."""
        params = []
        if campos:
            params += [('updateMask.fieldPaths', k) for k in campos]
        if update_time:
            params.append(('currentDocument.updateTime', update_time))
        body = {'fields': {k: enc(v) for k, v in data.items() if not k.startswith('_') and k != 'id' and (not campos or k in campos)}}
        r = self._req('PATCH', f'{self.base}/{path}', params=params, json=body)
        if r.status_code in (400, 409, 412) and update_time and ('FAILED_PRECONDITION' in r.text or r.status_code in (409, 412)):
            raise Conflicto(path)
        if r.status_code != 200: raise RuntimeError(f'Error guardando {path}: {r.status_code} {r.text[:300]}')
        return r.json().get('updateTime')

# ───────────────────────── reglas de la app (port de index.html) ─────────────────────────
EN_CURSO = ['Enviada', 'Validada', 'Exportada']
EST_SAP = ['', 'Reservado', 'Bloqueado', 'Desbloqueo solicitado', 'Desbloqueado', 'Cargado']
DEF_DISENO = ['aidrogo@tasa.com.pe', 'onavarro@tasa.com.pe', 'hcuellar@tasa.com.pe', 'lmontenegro@tasa.com.pe', 'gmunoz@tasa.com.pe']
CARGO_ALERTA = 'Encargado de diseño'
XL_HEAD = ['Código', 'Texto breve de material', 'Un. medida de entrada', 'Cantidad', 'Observaciones', 'Reserva', 'Posición', 'Estado',
           'Proyecto', 'Sistema', 'Grafo', 'Pos.', 'Empresa', 'Correo 1', 'Correo 2', 'Correo 3', 'Consideraciones', 'Fecha solicitud',
           'Fecha de carga', 'Nombre de actividad', 'Última Pos.']
XL_WCH = [10, 36, 8, 9, 14, 11, 8, 8, 16, 14, 10, 5, 12, 26, 28, 28, 40, 19, 11, 34, 9]

def tiene_res(it): return bool(it.get('reserva')) or bool(re.match(r'^reserv', str(it.get('estado') or ''), re.I))
def is_bloq(it): return it.get('estado') in ('Bloqueado', 'Desbloqueo solicitado', 'Desbloqueado') and not it.get('reserva')

def set_est_sap(it, v):
    b = it.get('bloqueo') or {}; it['bloqueo'] = b; n = now_iso(); it['estado'] = v
    if v == 'Bloqueado' and not b.get('desde'): b['desde'] = n
    if v in ('Cargado', 'Reservado'): b['cargado'] = n

def as_num(v):
    s = str(v if v is not None else '').strip()
    return int(s) if re.fullmatch(r'\d+', s) else (s or None)

def fmt(n):
    if n is None or n == '': return ''
    try: x = float(str(n).strip())
    except ValueError: return ''
    if math.isnan(x) or math.isinf(x): return ''
    x = math.floor(x * 1000 + 0.5) / 1000
    return str(int(x)) if x == int(x) else repr(x)

def dmy(d): return d.strftime('%d.%m.%Y')
def sanitize(s): return re.sub(r'[^A-Za-z0-9]', '_', str(s or ''))

class Datos:
    """Foto de la base de datos en un ciclo."""
    def __init__(self, fs):
        self.sol = fs.listar('solicitudes')
        self.emb = fs.listar('embarcaciones')
        self.empresas = {e['id']: e for e in fs.listar('empresas')}
        self.cuentas = {c['id']: c for c in fs.listar('cuentas')}
        d = fs.leer('config/diseno')
        self.diseno = (d or {}).get('list') or DEF_DISENO

    def emb_by(self, nombre):
        n = norm(nombre)
        return next((e for e in self.emb if norm(e.get('nombre')) == n), None)

    def ultima_pos(self, h):
        """Última Pos. del grafo según el listado de actividades (como ultimaPos() de la app)."""
        e = self.emb_by(h.get('proyecto'))
        if not e: return None
        g = re.sub(r'\D', '', str(h.get('grafo') or ''))
        sis = e.get('sistemas') or []
        s = next((x for x in sis if g and re.sub(r'\D', '', str(x.get('grafo') or '')) == g), None) or \
            next((x for x in sis if norm(x.get('nombre')) == norm(h.get('sistema'))), None)
        ps = []
        for a in (s or {}).get('actividades') or []:
            try: ps.append(int(str(a.get('pos')).strip()))
            except (TypeError, ValueError): pass
        return max(ps) if ps else None

    def cerrada(self, s):
        e = self.emb_by((s.get('header') or {}).get('proyecto'))
        return bool(e and e.get('estado') == 'Cerrado')

    def correos(self, s):
        c = self.cuentas.get(s.get('createdBy'))
        if not (c and c.get('empresaId') and c.get('empresaId') == s.get('empresaId')): c = None
        propios = split_mails(';'.join(str(c.get(k) or '') for k in ('correo1', 'correo2', 'correo3'))) if c else []
        e = self.empresas.get(s.get('empresaId')) or {}
        c1 = uniq(propios or split_mails(';'.join(str(e.get(k) or '') for k in ('correo1', 'correo2', 'correo3'))))
        emb = self.emb_by((s.get('header') or {}).get('proyecto'))
        con = [x for x in ((emb or {}).get('contactos') or []) if x.get('correo') and x.get('alerta') is not False]
        c3 = uniq(list(self.diseno) + [m for x in con if x.get('cargo') == CARGO_ALERTA for m in split_mails(x['correo'])])
        c2 = [m for m in uniq([m for x in con if x.get('cargo') != CARGO_ALERTA for m in split_mails(x['correo'])]) if m not in c3]
        return c1, c2, c3, e

def fecha_local(iso, tz):
    d = parse_iso(iso)
    return d.astimezone(tz).replace(tzinfo=None) if d else None

def nombre_archivo(s, tz):
    h = s.get('header') or {}
    d = fecha_local(s.get('fechaSolicitud'), tz) or datetime.now()
    txt = str(h.get('consideraciones') or h.get('nombreActividad') or '')[:52]
    return f"{sanitize(h.get('proyecto'))}-{sanitize(h.get('sistema'))}-{sanitize(txt)}_{d:%Y-%m-%d}_{d:%H-%M-%S}.xlsx"

def construir_excel(s, D, tz, destino):
    """Mismo contenido que buildXL() de la app (hoja 'Pedido')."""
    h = s.get('header') or {}
    c1, c2, c3, e = D.correos(s)
    fsol = fecha_local(s.get('fechaSolicitud'), tz)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Pedido'
    ws.append(XL_HEAD)
    for it in s.get('items') or []:
        est = it.get('estado') or ''
        ws.append([as_num(it.get('codigo')), it.get('desc'), it.get('um'), fmt(it.get('cantidad')), it.get('obs') or None,
                   as_num(it.get('reserva')), as_num(it.get('posicion')), None if est in EST_SAP else est,
                   h.get('proyecto'), h.get('sistema'), as_num(h.get('grafo')), as_num(h.get('pos')),
                   e.get('nombre') or h.get('empresaTxt') or '', ';'.join(c1), ';'.join(c2), ';'.join(c3),
                   h.get('consideraciones') or '', fsol, s.get('fechaCarga') or dmy(datetime.now(tz)), h.get('nombreActividad') or '',
                   D.ultima_pos(h)])
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 18).number_format = 'yyyy-mm-dd hh:mm:ss'
    for i, w in enumerate(XL_WCH, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    tmp = destino + '.tmp'
    wb.save(tmp); os.replace(tmp, destino)

# ───────────────────────── carpetas ─────────────────────────
def clave_proy(nombre):
    s = unicodedata.normalize('NFD', str(nombre or '')); s = ''.join(c for c in s if unicodedata.category(c) != 'Mn').upper()
    t = re.findall(r'[A-Z]+|\d+', s)
    t = ['RM' if x in ('REM', 'RM', 'REMOLCADOR') else x for x in t]
    return ' '.join(t)

def clave_sis(nombre):
    return norm(re.sub(r'^\s*\d+\s*[\.\-\)_]?\s*', '', str(nombre or '')))

class Carpetas:
    def __init__(self, cfg):
        self.c = cfg; self.proy = {}
        for raiz in cfg.raices:
            if not os.path.isdir(raiz): continue
            for d in os.listdir(raiz):
                p = os.path.join(raiz, d)
                if os.path.isdir(p): self.proy.setdefault(clave_proy(d), p)
        self.raices_ok = [r for r in cfg.raices if os.path.isdir(r)]

    def proyecto(self, nombre):
        a = self.c.alias.get(norm(nombre))
        if a:
            for raiz in self.raices_ok:
                p = os.path.join(raiz, a)
                if os.path.isdir(p): return p
            if os.path.isabs(a) and os.path.isdir(a): return a
        return self.proy.get(clave_proy(nombre))

    @staticmethod
    def sistema(pdir, sistema):
        if not pdir: return None
        k = clave_sis(sistema)
        for d in sorted(os.listdir(pdir)):
            p = os.path.join(pdir, d)
            if os.path.isdir(p) and clave_sis(d) == k: return p
        return None

    def subdir(self, sdir, nombre):
        for d in os.listdir(sdir):
            if norm(d) == norm(nombre) and os.path.isdir(os.path.join(sdir, d)): return os.path.join(sdir, d)
        return os.path.join(sdir, nombre)

    def exe(self, sdir):
        if os.path.isabs(self.c.exe_patron):  # instalación central del Reservatron (se le pasa --carpeta)
            return self.c.exe_patron if os.path.isfile(self.c.exe_patron) else None
        l = sorted(glob.glob(os.path.join(sdir, self.c.exe_patron)))
        return l[0] if l else None

    def carpetas_sistema(self):
        """Todas las carpetas de los sistemas habilitados, en todos los proyectos."""
        out = []
        for pdir in self.proy.values():
            for s in self.c.carpetas_esp():
                sd = self.sistema(pdir, s)
                if sd: out.append(sd)
        return out

# ───────────────────────── lectura del Excel devuelto ─────────────────────────
def leer_excel(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]; filas = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    if not filas: return []
    cab = [norm(x) for x in filas[0]]
    out = []
    for f in filas[1:]:
        if f is None or all(v is None or str(v).strip() == '' for v in f): continue
        r = {}
        for k, v in zip(cab, f):
            if k and k not in r: r[k] = cell_str(v)
        out.append(r)
    return out

def col_res(r):
    for k, v in r.items():
        if 'reserva' in k and not re.search(r'fecha|estado|tipo|clase', k) and str(v).strip() != '': return str(v).strip()
    return ''

def col_pos_res(r):
    for k in ('posicion', 'posicionreserva', 'posreserva', 'posres', 'itemreserva', 'posicionres'):
        if k in r and str(r[k]).strip() != '': return str(r[k]).strip()
    for k, v in r.items():
        if (k.startswith('posicion') or ('pos' in k and 'res' in k)) and str(v).strip() != '': return str(v).strip()
    return ''

def col_est(r):
    for k, v in r.items():
        if k.startswith('estado') and str(v).strip() != '': return str(v).strip()
    return ''

def parse_fecha(v):
    if isinstance(v, datetime): return v
    for f in ('%Y-%m-%d %H:%M:%S.%f', '%Y-%m-%d %H:%M:%S', '%d/%m/%Y %H:%M:%S', '%d/%m/%Y %H:%M', '%Y-%m-%dT%H:%M:%S.%f', '%Y-%m-%dT%H:%M:%S'):
        try: return datetime.strptime(str(v).strip(), f)
        except ValueError: pass
    return None

def codigo(v): return re.sub(r'\.0+$', '', str(v or '').strip())

def buscar_solicitud(rows, D, tz):
    r0 = rows[0]
    proy = r0.get('proyecto', ''); fsol = parse_fecha(r0.get('fechasolicitud'))
    cand = [s for s in D.sol if s.get('estado') != 'Anulada' and (not proy or norm((s.get('header') or {}).get('proyecto')) == norm(proy))]
    if fsol:
        m = [s for s in cand if (fl := fecha_local(s.get('fechaSolicitud'), tz)) and abs((fl - fsol).total_seconds()) < 2]
        if len(m) == 1: return m[0], 'fecha'
        if len(m) > 1: cand = m
    # respaldo: grafo + pos + códigos
    g, p = codigo(r0.get('grafo')), codigo(r0.get('pos'))
    codes = {codigo(r.get('codigo')) for r in rows if r.get('codigo')}
    mejor, ms = None, 0
    for s in cand:
        h = s.get('header') or {}
        if g and codigo(h.get('grafo')) != g: continue
        if p and codigo(h.get('pos')) != p: continue
        sc = {codigo(i.get('codigo')) for i in s.get('items') or [] if i.get('codigo')}
        if not sc or not codes: continue
        sim = len(sc & codes) / len(sc | codes)
        if sim > ms: mejor, ms = s, sim
    return (mejor, f'códigos {ms:.0%}') if mejor and ms >= 0.8 else (None, '')

def aplicar_reservas(s, rows):
    """Port de resXL(): devuelve (nR, nB, nSinNum, noEnc, cambió)."""
    items = s.get('items') or []; antes = json.dumps(items, sort_keys=True); est0 = s.get('estado')
    nR = nB = nSin = 0; used = set(); noenc = []; ahora = now_iso()
    for r in rows:
        c = codigo(r.get('codigo'))
        if not c: continue
        res, p, est = col_res(r), col_pos_res(r), col_est(r)
        bq = bool(re.search('bloq', est, re.I))
        i = next((ix for ix, it in enumerate(items) if ix not in used and str(it.get('codigo') or '').strip() == c and (not it.get('reserva') if bq else True)), -1)
        if i < 0: noenc.append(c); continue
        used.add(i); it = items[i]
        if bq:
            if not is_bloq(it): set_est_sap(it, 'Bloqueado')
            nB += 1; continue
        es_res = bool(re.search('reserv', est, re.I))
        if not res and not es_res: continue
        if (it.get('bloqueo') or {}).get('desde') or it.get('refId'):
            if not (it.get('reserva') == res and res): it['liberadoEn'] = ahora
        if is_bloq(it): set_est_sap(it, 'Reservado')
        if res: it['reserva'] = res; it['posicion'] = p or it.get('posicion') or ''
        else: nSin += 1
        it['estado'] = 'Reservado' if (es_res or not est) else est
        nR += 1
    if items and all(tiene_res(i) or is_bloq(i) for i in items) and any(tiene_res(i) for i in items) and s.get('estado') in EN_CURSO:
        s['estado'] = 'Reservada'
    elif s.get('estado') == 'Validada' and any(tiene_res(i) for i in items):
        s['estado'] = 'Exportada'
    s['items'] = items
    return nR, nB, nSin, noenc, (json.dumps(items, sort_keys=True) != antes or s.get('estado') != est0)

# ───────────────────────── el bot ─────────────────────────
class Bot:
    def __init__(self, cfg):
        self.c = cfg; self.fs = FS(cfg)
        self.est_path = os.path.join(BASE, 'estado_bot.json')
        try:
            with open(self.est_path, encoding='utf-8') as f: self.st = json.load(f)
        except Exception:
            self.st = {}
        self.st.setdefault('procesados', {})
        if not self.st.get('procesar_desde'):
            self.st['procesar_desde'] = cfg.procesar_desde or datetime.now().strftime('%Y-%m-%d %H:%M')
        if cfg.procesar_desde: self.st['procesar_desde'] = cfg.procesar_desde
        self.eventos = []; self.alertas = []; self.ultimo_error = ''
        self.pc = socket.gethostname(); self.win_user = getpass.getuser()

    def guardar_estado(self):
        if len(self.st['procesados']) > 3000:
            self.st['procesados'] = dict(sorted(self.st['procesados'].items(), key=lambda kv: kv[1].get('t', ''))[-2000:])
        tmp = self.est_path + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f: json.dump(self.st, f, ensure_ascii=False, indent=1)
        os.replace(tmp, self.est_path)

    def ev(self, nivel, m):
        getattr(log, {'ok': 'info', 'info': 'info', 'aviso': 'warning', 'error': 'error'}.get(nivel, 'info'))(m)
        self.eventos.append({'t': now_iso(), 'n': nivel, 'm': m}); self.eventos = self.eventos[-60:]

    def alerta(self, m):
        if m not in self.alertas:
            self.alertas.append(m); log.warning('ATENCIÓN: ' + m)

    def hist(self, s, a):
        s.setdefault('historial', []).append({'t': now_iso(), 'uid': self.c.usuario, 'a': a})

    def guardar_sol(self, s, campos, msg):
        if self.c.prueba:
            self.ev('info', '[PRUEBA] ' + msg); return True
        s['updatedAt'] = now_iso()
        try:
            self.fs.escribir('solicitudes/' + s['id'], s, campos=campos + ['updatedAt'], update_time=s.get('_ut'))
            return True
        except Conflicto:
            return False

    # ── 1. exportar ──
    def exportar(self, D, K):
        hechos = set()
        val = [s for s in D.sol if s.get('estado') == 'Validada']
        omit = lambda ref, m: (self.ev('info', f'Omitida {ref}: {m}') if self.c.prueba else log.debug(f'Omitida {ref}: {m}'))
        log.info(f'Solicitudes en la app: {len(D.sol)} · Validadas: {len(val)} · a automatizar ({self.c.desc}): '
                 f'{sum(1 for s in val if self.c.automatiza((s.get("header") or {}).get("sistema")))}')
        for s in val:
            h = s.get('header') or {}
            ref = f"{h.get('proyecto')} · {h.get('sistema') or '(sin sistema)'} · Pos {h.get('pos') or '—'}"
            if not self.c.automatiza(h.get('sistema')):
                omit(ref, f'el sistema "{h.get("sistema") or ""}" no está en "sistemas" de config.ini'); continue
            if D.cerrada(s):
                omit(ref, 'el proyecto está Cerrado'); continue
            if not h.get('grafo') or not h.get('pos'):
                self.alerta(f'{ref}: validada sin grafo o Pos.; no se exporta.'); continue
            if not s.get('items'):
                omit(ref, 'no tiene materiales'); continue
            pdir = K.proyecto(h.get('proyecto'))
            if not pdir:
                self.alerta(f'No encuentro la carpeta del proyecto "{h.get("proyecto")}" en {"; ".join(self.c.raices)}. Agregue un alias en config.ini.'); continue
            carp = self.c.carpeta_de(h.get('sistema'))
            sdir = K.sistema(pdir, carp)
            if not sdir:
                self.alerta(f'No encuentro la carpeta "{carp}" dentro de {pdir}.'); continue
            sol_dir = K.subdir(sdir, self.c.sub_sol); res_dir = K.subdir(sdir, self.c.sub_res)
            nombre = nombre_archivo(s, self.c.tz); destino = os.path.join(sol_dir, nombre)
            ya = os.path.exists(destino) or os.path.exists(os.path.join(res_dir, nombre))
            if self.c.prueba:
                self.ev('info', f'[PRUEBA] Exportaría {ref} → {destino}{" (ya existe)" if ya else ""}'); continue
            for intento in range(2):
                if not ya:
                    os.makedirs(sol_dir, exist_ok=True)
                    construir_excel(s, D, self.c.tz, destino)
                if not s.get('fechaCarga'): s['fechaCarga'] = dmy(datetime.now(self.c.tz))
                s['estado'] = 'Exportada'
                self.hist(s, f'🤖 Bot: {"el Excel ya estaba" if ya else "generó el Excel"} en {sol_dir if not ya else "la carpeta"} ({nombre})')
                if self.guardar_sol(s, ['estado', 'fechaCarga', 'historial'], ''):
                    self.ev('ok', f'Exportada {ref} → {destino}'); hechos.add(sdir)
                    self.st.setdefault('exportados', {})[destino] = now_iso(); break
                # conflicto: alguien la editó; volver a leer
                s2 = self.fs.leer('solicitudes/' + s['id'])
                if not s2 or s2.get('estado') != 'Validada':
                    if not ya and os.path.exists(destino): os.remove(destino)
                    self.ev('aviso', f'{ref}: cambió en la app mientras se exportaba; se omitió.'); break
                s.clear(); s.update(s2)
        return hechos

    # ── 2. ejecutable ──
    def ejecutar(self, carpetas, K):
        for sdir in sorted(carpetas):
            sol_dir = K.subdir(sdir, self.c.sub_sol)
            pend = glob.glob(os.path.join(sol_dir, '*.xls*'))
            pend = [p for p in pend if not os.path.basename(p).startswith('~$')]
            if not pend: continue
            exe = K.exe(sdir)
            if not exe:
                self.alerta(f'No hay ejecutable ({self.c.exe_patron}) en {sdir}.'); continue
            if self.c.prueba:
                self.ev('info', f'[PRUEBA] Ejecutaría {exe} ({len(pend)} archivo(s) en Solicitados)'); continue
            self.ev('info', f'Ejecutando {os.path.basename(exe)} en {sdir} ({len(pend)} archivo(s))…')
            self.estado_app('ejecutando', f'Cargando en SAP: {os.path.basename(os.path.dirname(sdir))} / {os.path.basename(sdir)}')
            t0 = time.time()
            try:
                kw = {'creationflags': subprocess.CREATE_NEW_CONSOLE} if os.name == 'nt' else {}
                cmd = ([exe] if exe.lower().endswith('.exe') or os.name == 'nt' else [sys.executable, exe]) + self.c.exe_args
                if os.path.normcase(os.path.dirname(os.path.abspath(exe))) != os.path.normcase(os.path.abspath(sdir)):
                    cmd += ['--carpeta', sdir]
                self.st.setdefault('intentos', {})[sdir] = now_iso()
                import threading
                fin_hb = threading.Event()
                def _latido():  # mantiene "en línea" en la app mientras el Reservatron trabaja
                    while not fin_hb.wait(120):
                        self.estado_app('ejecutando', f'Cargando en SAP: {os.path.basename(os.path.dirname(sdir))} / {os.path.basename(sdir)}')
                threading.Thread(target=_latido, daemon=True).start()
                try:
                    r = subprocess.run(cmd, cwd=sdir, timeout=self.c.exe_timeout, stdin=subprocess.DEVNULL, **kw)
                finally:
                    fin_hb.set()
                quedan = [p for p in pend if os.path.exists(p)]
                resumen = self.resumen_reservatron(sdir)
                proy = os.path.basename(os.path.dirname(sdir))
                txt = {0: 'todo procesado', 1: 'quedaron archivos pendientes', 2: 'error general (¿SAP abierto y con sesión?)',
                       3: 'otro Reservatron ya estaba trabajando en esa carpeta'}.get(r.returncode, f'código {r.returncode}')
                lvl = 'ok' if r.returncode == 0 and not quedan else ('error' if r.returncode == 2 else 'aviso')
                self.ev(lvl, f'Reservatron en {proy}: {txt} ({time.time() - t0:.0f} s). {len(pend) - len(quedan)} de {len(pend)} archivo(s) pasaron a Reservados.'
                             + (f' {resumen}' if resumen else ''))
                if r.returncode == 2: self.alerta(f'Reservatron en {proy} no pudo trabajar: {resumen or "revise SAP"}')
                for q in quedan: self.alerta(f'Quedó en Solicitados sin procesar: {q}')
            except subprocess.TimeoutExpired:
                self.ev('error', f'{os.path.basename(exe)} superó {self.c.exe_timeout // 60} min en {sdir}; se dejó de esperar.')
            except Exception as e:
                self.ev('error', f'No se pudo ejecutar {exe}: {e}')

    @staticmethod
    def resumen_reservatron(sdir):
        """Última línea RESUMEN del Reservatron_auto.log de esa carpeta."""
        try:
            with open(os.path.join(sdir, 'Reservatron_auto.log'), encoding='utf-8', errors='replace') as f:
                lineas = f.readlines()[-60:]
            for l in reversed(lineas):
                if l.startswith('RESUMEN:'): return l.strip()
        except Exception:
            pass
        return ''

    def carpetas_a_ejecutar(self, exportadas, K, forzar):
        """Carpetas donde correr el Reservatron: las que recibieron Excel en este ciclo y, para reintentar,
        las que aún tienen en Solicitados algún Excel que dejó el bot (nunca por archivos ajenos)."""
        exp = self.st.setdefault('exportados', {})
        for pth in list(exp):
            if not os.path.exists(pth): del exp[pth]
        carp = set(exportadas)
        ahora = datetime.now(timezone.utc)
        for pth in exp:
            sdir = os.path.dirname(os.path.dirname(pth))
            ult = parse_iso(self.st.get('intentos', {}).get(sdir))
            if forzar or (self.c.reintentar > 0 and (not ult or (ahora - ult).total_seconds() > self.c.reintentar * 60)):
                carp.add(sdir)
        return carp

    # ── 3. reservados ──
    def leer_reservados(self, D, K):
        desde = parse_fecha(self.st['procesar_desde'] + (':00' if len(self.st['procesar_desde']) == 16 else '')) or parse_fecha(self.st['procesar_desde'] + ' 00:00:00')
        n = 0
        for sdir in K.carpetas_sistema():
            rdir = K.subdir(sdir, self.c.sub_res)
            if not os.path.isdir(rdir): continue
            for path in sorted(glob.glob(os.path.join(rdir, '*.xls*'))):
                base = os.path.basename(path)
                if base.startswith('~$') or base.endswith('.tmp'): continue
                try: stt = os.stat(path)
                except OSError: continue
                key = f'{path}|{int(stt.st_mtime)}|{stt.st_size}'
                if key in self.st['procesados']: continue
                if desde and datetime.fromtimestamp(stt.st_mtime) < desde:
                    continue
                if time.time() - stt.st_mtime < ESPERA_COPIA: continue  # aún se está copiando
                try:
                    rows = leer_excel(path)
                except PermissionError:
                    continue
                except Exception as e:
                    self.ev('aviso', f'No se pudo leer {base}: {e}'); self.st['procesados'][key] = {'t': now_iso(), 'r': 'ilegible'}; continue
                if not rows:
                    self.st['procesados'][key] = {'t': now_iso(), 'r': 'vacío'}; continue
                s, como = buscar_solicitud(rows, D, self.c.tz)
                if not s:
                    self.alerta(f'Excel sin solicitud en la app: {path}')
                    self.st['procesados'][key] = {'t': now_iso(), 'r': 'sin solicitud'}; continue
                h = s.get('header') or {}; ref = f"{h.get('proyecto')} · {h.get('sistema')} · Pos {h.get('pos') or '—'}"
                ok = False
                for intento in range(3):
                    nR, nB, nSin, noenc, cambio = aplicar_reservas(s, rows)
                    if not cambio:
                        ok = True; break
                    det = f'{nR} reservado(s)' + (f', {nB} bloqueado(s)' if nB else '') + (f', {nSin} sin número' if nSin else '') + (f', {len(noenc)} código(s) no encontrados' if noenc else '')
                    self.hist(s, f'🤖 Bot: actualizó reserva desde {base} ({det})')
                    if self.guardar_sol(s, ['items', 'estado', 'historial'], f'Actualizaría {ref} desde {base}: {det}'):
                        if not self.c.prueba:
                            self.ev('ok', f'Reserva registrada: {ref} ← {base} ({det}; asociada por {como}){" → Reservada" if s.get("estado") == "Reservada" else ""}')
                        if noenc: self.alerta(f'{base}: códigos que no están en la solicitud: {", ".join(noenc[:8])}')
                        ok = True; n += 1; break
                    s2 = self.fs.leer('solicitudes/' + s['id'])
                    if not s2: break
                    s.clear(); s.update(s2)
                if ok and not self.c.prueba:
                    self.st['procesados'][key] = {'t': now_iso(), 'r': s['id']}
        return n

    # ── estado en la app ──
    def estado_app(self, fase, detalle=''):
        if self.c.prueba: return
        data = {'version': VERSION, 'pc': self.pc, 'usuarioWindows': self.win_user, 'ultimoContacto': now_iso(),
                'fase': fase, 'detalle': detalle, 'intervaloMin': self.c.intervalo, 'sistemas': [self.c.desc],
                'raices': self.c.raices, 'ejecutaExe': self.c.ejecutar_exe, 'ultimoCiclo': self.st.get('ultimo_ciclo', ''),
                'proximoCiclo': self.st.get('proximo_ciclo', ''), 'ultimoError': self.ultimo_error,
                'log': self.eventos[-40:], 'alertas': self.alertas[:30], 'resumen': self.st.get('resumen', {})}
        try:
            self.fs.escribir('botEstado/estado', data)
        except Exception as e:
            log.warning(f'No se pudo publicar el estado en la app: {e}')

    def orden_pendiente(self):
        try:
            o = self.fs.leer('botEstado/orden')
        except Exception:
            return False
        p = (o or {}).get('pedido') or ''
        if p and p > self.st.get('ultima_orden', ''):
            self.st['ultima_orden'] = p
            self.ev('info', f'"Ejecutar ahora" solicitado desde la app por {(o or {}).get("por") or "?"}.')
            return True
        return False

    # ── ciclo completo ──
    def ciclo(self, forzar_exe=False):
        self.alertas = []
        t0 = time.time()
        try:
            self.estado_app('trabajando', 'Revisando solicitudes y carpetas…')
            K = Carpetas(self.c)
            if not K.raices_ok:
                raise RuntimeError(f'No se puede acceder a la carpeta de reservas ({"; ".join(self.c.raices)}). ¿Está conectada la unidad Z:?')
            D = Datos(self.fs)
            exp = self.exportar(D, K) if self.c.exportar else set()
            if self.c.ejecutar_exe:
                carp = self.carpetas_a_ejecutar(exp, K, forzar_exe)
                self.ejecutar(carp, K)
                if carp: D = Datos(self.fs)
            nres = self.leer_reservados(D, K) if self.c.leer_res else 0
            pend = []
            for sdir in K.carpetas_sistema():
                sd = K.subdir(sdir, self.c.sub_sol)
                for p in glob.glob(os.path.join(sd, '*.xls*')):
                    if not os.path.basename(p).startswith('~$') and time.time() - os.path.getmtime(p) > 3600: pend.append(p)
            for p in pend[:10]: self.alerta(f'En Solicitados hace más de 1 h: {p}')
            self.ultimo_error = ''
            self.st['ultimo_ciclo'] = now_iso()
            self.st['resumen'] = {'exportadas': len(exp), 'reservasLeidas': nres, 'proyectos': len(K.proy), 'seg': round(time.time() - t0)}
            log.info(f'Ciclo terminado en {time.time() - t0:.0f} s: {len(exp)} carpeta(s) con exportaciones, {nres} reserva(s) leída(s).')
        except Exception as e:
            self.ultimo_error = f'{type(e).__name__}: {e}'
            self.ev('error', f'Error en el ciclo: {e}'); log.debug(traceback.format_exc())
        finally:
            self.st['proximo_ciclo'] = (datetime.now(timezone.utc) + timedelta(minutes=self.c.intervalo)).isoformat(timespec='seconds').replace('+00:00', 'Z')
            if not self.c.prueba: self.guardar_estado()
            self.estado_app('esperando', '')

    def correr(self, una_vez=False):
        self.ev('info', f'Bot de reservas v{VERSION} iniciado en {self.pc} ({self.win_user}). Automatiza: {self.c.desc}. Cada {self.c.intervalo} min.'
                        + (' MODO PRUEBA: no se escribe nada.' if self.c.prueba else ''))
        self.orden_pendiente()  # descarta órdenes antiguas
        while True:
            self.ciclo()
            if una_vez: return
            fin = time.time() + self.c.intervalo * 60; ult_hb = time.time()
            while time.time() < fin:
                time.sleep(30)
                if self.orden_pendiente():
                    self.ciclo(forzar_exe=True); fin = time.time() + self.c.intervalo * 60
                if time.time() - ult_hb > 300:
                    self.estado_app('esperando', ''); ult_hb = time.time()

def diagnostico(cfg):
    print(f'\n=== Diagnóstico del bot (no escribe nada) ===\nUsuario: {cfg.usuario}@{cfg.dominio} · Proyecto Firebase: {cfg.project}')
    fs = FS(cfg); D = Datos(fs); K = Carpetas(cfg)
    print(f'\nRaíces: ' + '; '.join(f'{r} [{"OK" if os.path.isdir(r) else "NO EXISTE"}]' for r in cfg.raices))
    print(f'Carpetas de proyecto encontradas: {len(K.proy)}')
    from collections import Counter
    print(f'\nSolicitudes que ve el bot: {len(D.sol)}  ' + ', '.join(f'{k}: {v}' for k, v in Counter(s.get("estado") for s in D.sol).most_common()))
    for s in D.sol:
        if s.get('estado') not in ('Validada', 'Exportada', 'Enviada'): continue
        h = s.get('header') or {}
        pdir = K.proyecto(h.get('proyecto')); sdir = K.sistema(pdir, cfg.carpeta_de(h.get('sistema'))) if pdir else None
        print(f'\n• [{s.get("estado")}] {h.get("proyecto")} · sistema "{h.get("sistema")}" · grafo {h.get("grafo") or "—"} · Pos {h.get("pos") or "—"} · {len(s.get("items") or [])} materiales')
        print(f'    se automatiza: {"SÍ" if cfg.automatiza(h.get("sistema")) else "NO"} · carpeta proyecto: {pdir or "NO ENCONTRADA"}'
              f' · carpeta destino: {sdir or "NO ENCONTRADA (" + cfg.carpeta_de(h.get("sistema")) + ")"}')
    print()

def main():
    ap = argparse.ArgumentParser(description='Bot de reservas — Pedidos de Material TASA')
    ap.add_argument('--config', default=os.path.join(BASE, 'config.ini'))
    ap.add_argument('--una-vez', action='store_true', help='Un solo ciclo y termina')
    ap.add_argument('--prueba', action='store_true', help='No escribe en la app ni en carpetas, no ejecuta el .exe')
    ap.add_argument('--diagnostico', action='store_true', help='Lista las solicitudes que ve el bot y las carpetas que encuentra')
    a = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)-7s %(message)s', datefmt='%Y-%m-%d %H:%M:%S',
                        handlers=[logging.StreamHandler(sys.stdout),
                                  RotatingFileHandler(os.path.join(BASE, 'bot_reservas.log'), maxBytes=2_000_000, backupCount=3, encoding='utf-8')])
    try:
        cfg = Cfg(a.config, prueba=a.prueba or a.diagnostico)
        if a.diagnostico:
            diagnostico(cfg); return
        Bot(cfg).correr(una_vez=a.una_vez)
    except KeyboardInterrupt:
        log.info('Bot detenido por el usuario.')
    except SystemExit as e:
        log.error(str(e));
        if not a.una_vez: input('Enter para salir…')
        sys.exit(1)

if __name__ == '__main__':
    main()
