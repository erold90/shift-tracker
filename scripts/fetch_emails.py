#!/usr/bin/env python3
"""
Script per estrarre email di notifica turni da Gmail e parsare i servizi.
Utilizza Gmail API per l'autenticazione sicura.

Gestisce:
- Turni di servizio (PRESENZA, etc.)
- Eliminazioni turni
- Domande di licenza (ordinaria, straordinaria, speciale)
- Aggiornamenti: l'email più recente per lo stesso giorno sovrascrive le precedenti
- Calcolo ore straordinarie (> 6 ore giornaliere)
"""

import os
import json
import re
import base64
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, List, Dict, Any
from dataclasses import dataclass, asdict, field
from collections import defaultdict
import calendar

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Scope necessari per leggere le email
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

# Path dei file
BASE_DIR = Path(__file__).parent.parent
CREDENTIALS_FILE = BASE_DIR / 'credentials.json'
TOKEN_FILE = BASE_DIR / 'token.json'
DATA_FILE = BASE_DIR / 'data' / 'servizi.json'

# Mittente delle email di notifica turni (configurabile via variabile d'ambiente)
MEMO_SENDER = os.environ.get('MEMO_SENDER', 'noreply@example.com')

# Ore standard di un turno (per calcolo straordinario)
ORE_TURNO_STANDARD = 6.0

# ID dipendente (configurabile via variabile d'ambiente)
EMPLOYEE_ID = os.environ.get('EMPLOYEE_ID', '000000')


def sanitize_dettaglio(dettaglio: str) -> str:
    """Rimuove riferimenti identificativi dal campo dettaglio."""
    if not dettaglio:
        return dettaglio

    # Sostituzioni per generalizzare - ordine importante!
    replacements = [
        # Riferimenti legali e sigle
        (r'\s*\(ex art\.\d+\s*L\.\s*\d+/\d+\s*-\s*U\.C\.I\.S\.', ''),
        (r'U\.?C\.?I\.?S\.?', ''),

        # Servizi specifici militari/polizia
        (r'Militare servizio caserma/addetto ricezione pubblico.*', 'Servizio interno'),
        (r'Militare servizio caserma.*', 'Servizio interno'),
        (r'servizio caserma.*', 'Servizio interno'),
        (r'Scorta a persona.*', 'Servizio esterno'),
        (r'Indagini di Polizia Giudiziaria', 'Attività operativa'),
        (r'Polizia Giudiziaria', 'Attività operativa'),
        (r'Accompagnamento a collaboratore di giustizia.*', 'Servizio esterno'),
        (r'collaboratore di giustizia', 'soggetto protetto'),
        (r'Testimonianza per fatti inerenti al servizio', 'Impegno istituzionale'),
        (r'Testimonianza.*', 'Impegno istituzionale'),

        # Termini militari generici
        (r'\bMilitare\b', 'Operatore'),
        (r'\bmilitare\b', 'operatore'),
        (r'\bCarabinieri\b', 'Ente'),
        (r'\bcarabinieri\b', 'ente'),
        (r'\bArma\b', 'Ente'),
        (r'\bcaserma\b', 'sede'),
        (r'\bCaserma\b', 'Sede'),
    ]

    result = dettaglio
    for pattern, replacement in replacements:
        result = re.sub(pattern, replacement, result, flags=re.IGNORECASE)

    return result.strip()


@dataclass
class Turno:
    """Rappresenta un singolo turno di servizio."""
    id: str
    tipo: str  # PRESENZA, REPERIBILITA, etc.
    dettaglio: str
    matricola: str
    data: str  # YYYY-MM-DD
    ora_inizio: str  # HH:MM
    ora_fine: str  # HH:MM
    durata_ore: float
    is_straordinario: bool
    ore_ordinarie: float
    ore_straordinario: float
    email_date: str  # Timestamp email per determinare quale è più recente
    email_id: str
    stato: str  # attivo, eliminato

    def to_dict(self):
        # Escludi matricola e sanitizza dettaglio per privacy
        d = asdict(self)
        d.pop('matricola', None)
        d['dettaglio'] = sanitize_dettaglio(d.get('dettaglio', ''))
        return d


@dataclass
class Licenza:
    """Rappresenta una domanda di licenza."""
    id: str
    tipo: str  # ordinaria, straordinaria, speciale, riposo_donatori
    motivo: str  # RIPOSO MEDICO, recupero festività, etc.
    stato: str  # Presentata, Validata, Approvata, Annullata
    data_inizio: str
    data_fine: str
    email_date: str
    email_id: str

    def to_dict(self):
        return asdict(self)


@dataclass
class Giornata:
    """Rappresenta tutti i turni di una giornata."""
    data: str
    turni: List[Turno] = field(default_factory=list)
    ore_totali: float = 0.0
    ore_ordinarie: float = 0.0
    ore_straordinario: float = 0.0
    is_licenza: bool = False
    tipo_licenza: str = ""

    def to_dict(self):
        return {
            'data': self.data,
            'turni': [t.to_dict() for t in self.turni],
            'ore_totali': self.ore_totali,
            'ore_ordinarie': self.ore_ordinarie,
            'ore_straordinario': self.ore_straordinario,
            'is_licenza': self.is_licenza,
            'tipo_licenza': self.tipo_licenza
        }


# Mappa mesi italiani
MESI = {
    'gennaio': '01', 'febbraio': '02', 'marzo': '03', 'aprile': '04',
    'maggio': '05', 'giugno': '06', 'luglio': '07', 'agosto': '08',
    'settembre': '09', 'ottobre': '10', 'novembre': '11', 'dicembre': '12'
}


def get_festivi_italiani(anno: int) -> Dict[str, str]:
    """
    Restituisce i giorni festivi italiani per l'anno specificato.
    Chiave: data in formato YYYY-MM-DD
    Valore: nome della festività
    """
    festivi = {
        f"{anno}-01-01": "Capodanno",
        f"{anno}-01-06": "Epifania",
        f"{anno}-04-25": "Festa della Liberazione",
        f"{anno}-05-01": "Festa dei Lavoratori",
        f"{anno}-06-02": "Festa della Repubblica",
        f"{anno}-08-15": "Ferragosto",
        f"{anno}-11-01": "Tutti i Santi",
        f"{anno}-12-08": "Immacolata Concezione",
        f"{anno}-12-25": "Natale",
        f"{anno}-12-26": "Santo Stefano",
    }

    # Calcola Pasqua (algoritmo di Gauss)
    # Per il 2025: Pasqua è il 20 aprile
    if anno == 2025:
        pasqua = datetime(2025, 4, 20)
    else:
        # Algoritmo generico per altri anni
        a = anno % 19
        b = anno // 100
        c = anno % 100
        d = b // 4
        e = b % 4
        f = (b + 8) // 25
        g = (b - f + 1) // 3
        h = (19 * a + b - d - g + 15) % 30
        i = c // 4
        k = c % 4
        l = (32 + 2 * e + 2 * i - h - k) % 7
        m = (a + 11 * h + 22 * l) // 451
        mese = (h + l - 7 * m + 114) // 31
        giorno = ((h + l - 7 * m + 114) % 31) + 1
        pasqua = datetime(anno, mese, giorno)

    # Pasqua e Lunedì dell'Angelo
    festivi[pasqua.strftime("%Y-%m-%d")] = "Pasqua"
    lunedi_angelo = pasqua + timedelta(days=1)
    festivi[lunedi_angelo.strftime("%Y-%m-%d")] = "Lunedì dell'Angelo"

    return festivi


def get_all_sundays(anno: int, mese_inizio: int = 1, mese_fine: int = 12) -> List[str]:
    """Restituisce tutte le domeniche dell'anno (o range di mesi)."""
    domeniche = []
    for mese in range(mese_inizio, mese_fine + 1):
        cal = calendar.Calendar()
        for giorno in cal.itermonthdays2(anno, mese):
            if giorno[0] != 0 and giorno[1] == 6:  # 6 = domenica
                domeniche.append(f"{anno}-{str(mese).zfill(2)}-{str(giorno[0]).zfill(2)}")
    return domeniche


def get_gmail_service():
    """Ottiene il servizio Gmail autenticato."""
    creds = None

    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_FILE.exists():
                raise FileNotFoundError(
                    f"File credentials.json non trovato in {CREDENTIALS_FILE}. "
                    "Scaricalo da Google Cloud Console."
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE), SCOPES
            )
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)


def strip_html_tags(html: str) -> str:
    """Rimuove i tag HTML e converte in testo."""
    # Sostituisci <br> con newline
    text = re.sub(r'<br\s*/?>', '\n', html, flags=re.IGNORECASE)
    # Rimuovi tutti gli altri tag HTML
    text = re.sub(r'<[^>]+>', '', text)
    # Decodifica entità HTML comuni
    text = text.replace('&nbsp;', ' ')
    text = text.replace('&amp;', '&')
    text = text.replace('&lt;', '<')
    text = text.replace('&gt;', '>')
    text = text.replace('&quot;', '"')
    return text


def get_email_body(msg) -> str:
    """Estrae il corpo testuale da un messaggio Gmail."""
    body = ""

    # Prima prova text/plain
    if 'parts' in msg['payload']:
        for part in msg['payload']['parts']:
            if part['mimeType'] == 'text/plain':
                if 'data' in part['body']:
                    body = base64.urlsafe_b64decode(
                        part['body']['data']
                    ).decode('utf-8', errors='ignore')
                    break
        # Se non trovato text/plain, prova text/html
        if not body:
            for part in msg['payload']['parts']:
                if part['mimeType'] == 'text/html':
                    if 'data' in part['body']:
                        html = base64.urlsafe_b64decode(
                            part['body']['data']
                        ).decode('utf-8', errors='ignore')
                        body = strip_html_tags(html)
                        break
    elif 'body' in msg['payload'] and 'data' in msg['payload']['body']:
        raw = base64.urlsafe_b64decode(
            msg['payload']['body']['data']
        ).decode('utf-8', errors='ignore')
        # Controlla se è HTML
        if '<br' in raw.lower() or '<b>' in raw.lower():
            body = strip_html_tags(raw)
        else:
            body = raw

    return body


def get_email_date(msg) -> str:
    """Estrae la data di ricezione dell'email come timestamp ISO."""
    headers = msg['payload'].get('headers', [])
    for header in headers:
        if header['name'].lower() == 'date':
            try:
                date_str = header['value']
                date_str = re.sub(r'\s*\([^)]+\)\s*$', '', date_str)
                date_str = re.sub(r'\s*[+-]\d{4}\s*$', '', date_str)

                for fmt in [
                    '%a, %d %b %Y %H:%M:%S',
                    '%d %b %Y %H:%M:%S',
                    '%a %d %b %Y %H:%M:%S'
                ]:
                    try:
                        dt = datetime.strptime(date_str.strip(), fmt)
                        return dt.isoformat()
                    except:
                        continue
            except:
                pass
    return datetime.now().isoformat()


def get_email_subject(msg) -> str:
    """Estrae l'oggetto dell'email."""
    headers = msg['payload'].get('headers', [])
    for header in headers:
        if header['name'].lower() == 'subject':
            return header['value']
    return ""


def parse_data_italiana(giorno: str, mese: str, anno: str) -> str:
    """Converte data italiana in formato YYYY-MM-DD."""
    mese_num = MESI.get(mese.lower(), '01')
    return f"{anno}-{mese_num}-{giorno.zfill(2)}"


def parse_turno_servizio(body: str, email_date: str, msg_id: str, subject: str) -> Optional[Turno]:
    """
    Parsa un'email di tipo "Aggiornamento turno di servizio".

    Formato:
    Servizio di PRESENZA (Tipo servizio):
    Matricola impiegato: XXXXXX
    Inizio: ore 14:00 del giorno 15/dicembre/2025
    Fine: ore 17:30 del giorno 15/dicembre/2025
    """
    try:
        # Pattern per il tipo di servizio
        tipo_match = re.search(
            r'Servizio di\s+(\w+)\s*\(([^)]+)\)',
            body,
            re.IGNORECASE
        )

        # Pattern per matricola (generico)
        matricola_match = re.search(
            r'Matricola[^:]*:\s*(\d+)',
            body
        )

        # Pattern per data/ora inizio
        inizio_match = re.search(
            r'Inizio:\s*ore\s*(\d{1,2}[:.]\d{2})\s*del giorno\s*(\d{1,2})/(\w+)/(\d{4})',
            body,
            re.IGNORECASE
        )

        # Pattern per data/ora fine (opzionale - alcune email non ce l'hanno)
        fine_match = re.search(
            r'Fine:\s*ore\s*(\d{1,2}[:.]\d{2})\s*del giorno\s*(\d{1,2})/(\w+)/(\d{4})',
            body,
            re.IGNORECASE
        )

        # Verifica campi obbligatori
        if not all([tipo_match, matricola_match, inizio_match]):
            return None

        # Se manca "Fine", ignora l'email (incompleta)
        if not fine_match:
            print(f"  [SKIP] Email {msg_id}: manca orario Fine (email incompleta)")
            return None

        tipo_servizio = tipo_match.group(1).strip().upper()
        dettaglio = tipo_match.group(2).strip()
        matricola = matricola_match.group(1)

        # Parse inizio
        ora_inizio = inizio_match.group(1).replace('.', ':')
        data = parse_data_italiana(
            inizio_match.group(2),
            inizio_match.group(3),
            inizio_match.group(4)
        )

        # Parse fine
        ora_fine = fine_match.group(1).replace('.', ':')
        data_fine = parse_data_italiana(
            fine_match.group(2),
            fine_match.group(3),
            fine_match.group(4)
        )

        # Calcola durata
        dt_inizio = datetime.strptime(f"{data} {ora_inizio}", "%Y-%m-%d %H:%M")
        dt_fine = datetime.strptime(f"{data_fine} {ora_fine}", "%Y-%m-%d %H:%M")
        durata = (dt_fine - dt_inizio).total_seconds() / 3600

        if durata < 0:
            durata = 0

        # Calcola ore ordinarie vs straordinario
        ore_ordinarie = min(durata, ORE_TURNO_STANDARD)
        ore_straordinario = max(0, durata - ORE_TURNO_STANDARD)
        is_straordinario = ore_straordinario > 0

        return Turno(
            id=f"{data}_{ora_inizio.replace(':', '')}_{msg_id[:8]}",
            tipo=tipo_servizio,
            dettaglio=dettaglio,
            matricola=matricola,
            data=data,
            ora_inizio=ora_inizio,
            ora_fine=ora_fine,
            durata_ore=round(durata, 2),
            is_straordinario=is_straordinario,
            ore_ordinarie=round(ore_ordinarie, 2),
            ore_straordinario=round(ore_straordinario, 2),
            email_date=email_date,
            email_id=msg_id,
            stato='attivo'
        )

    except Exception as e:
        print(f"  [ERROR] Parsing turno {msg_id}: {e}")
        return None


def parse_eliminazione_turno(body: str, email_date: str, msg_id: str, subject: str) -> Optional[Dict]:
    """
    Parsa un'email di tipo "Eliminazione turno pianificato".
    Restituisce info per marcare il turno come eliminato.

    Formato body:
    E' appena stato eliminato il seguente servizio:
    Servizio di PRESENZA (Tipo servizio):
    Matricola impiegato: XXXXXX
    Inizio: ore 14:00 del giorno 16/dicembre/2025
    Fine: ore 20:00 del giorno 16/dicembre/2025
    """
    try:
        # Prima prova a estrarre dal body (più preciso)
        inizio_match = re.search(
            r'Inizio:\s*ore\s*(\d{1,2}[:.]\d{2})\s*del giorno\s*(\d{1,2})/(\w+)/(\d{4})',
            body,
            re.IGNORECASE
        )

        if inizio_match:
            ora_inizio = inizio_match.group(1).replace('.', ':')
            data = parse_data_italiana(
                inizio_match.group(2),
                inizio_match.group(3),
                inizio_match.group(4)
            )

            # Estrai anche ora fine se presente
            fine_match = re.search(
                r'Fine:\s*ore\s*(\d{1,2}[:.]\d{2})\s*del giorno\s*(\d{1,2})/(\w+)/(\d{4})',
                body,
                re.IGNORECASE
            )
            ora_fine = fine_match.group(1).replace('.', ':') if fine_match else ""

            return {
                'tipo': 'eliminazione',
                'data': data,
                'ora_inizio': ora_inizio,
                'ora_fine': ora_fine,
                'email_date': email_date,
                'email_id': msg_id
            }

        # Fallback: estrai data dall'oggetto
        data_match = re.search(
            r'Eliminazione turno pianificato per il giorno\s*(\d{1,2})/(\d{1,2})/(\d{4})',
            subject
        )

        if data_match:
            giorno = data_match.group(1).zfill(2)
            mese = data_match.group(2).zfill(2)
            anno = data_match.group(3)
            data = f"{anno}-{mese}-{giorno}"

            return {
                'tipo': 'eliminazione',
                'data': data,
                'ora_inizio': '',
                'ora_fine': '',
                'email_date': email_date,
                'email_id': msg_id
            }

        return None

    except Exception as e:
        print(f"  [ERROR] Parsing eliminazione {msg_id}: {e}")
        return None


def parse_licenza(body: str, email_date: str, msg_id: str, subject: str) -> Optional[Licenza]:
    """
    Parsa un'email di domanda di licenza.

    Formato body:
    Domanda di Licenza ordinaria in stato Approvata.
    Data inizio: 06/dicembre/2025
    Data fine: 06/dicembre/2025
    Tipo fruizione: INTERA GIORNATA

    Tipi:
    - Domanda di Licenza ordinaria
    - Domanda di Licenza straordinaria per gravi motivi (RIPOSO MEDICO)
    - Domanda di Licenza speciale (recupero festività soppresse)
    - Domanda di Riposo per donatori di sangue ed emocomponenti in struttura civile
    """
    try:
        # Determina tipo licenza e motivo
        tipo_licenza = "ordinaria"
        motivo = ""
        stato = ""
        tipo_fruizione = ""

        subject_lower = subject.lower()
        body_lower = body.lower()

        if "straordinaria" in subject_lower:
            tipo_licenza = "straordinaria"
            motivo_match = re.search(r'\(([^)]+)\)', subject)
            if motivo_match:
                motivo = motivo_match.group(1)
        elif "speciale" in subject_lower:
            tipo_licenza = "speciale"
            motivo_match = re.search(r'\(([^)]+)\)', subject)
            if motivo_match:
                motivo = motivo_match.group(1)
        elif "riposo per donatori" in subject_lower:
            tipo_licenza = "riposo_donatori"
            motivo = "donazione sangue"

        # Estrai stato dall'oggetto
        stato_match = re.search(
            r'in stato\s+(\w+)',
            subject,
            re.IGNORECASE
        )
        if stato_match:
            stato = stato_match.group(1)

        # Estrai date dal body
        # Formato: Data inizio: 06/dicembre/2025
        data_inizio = ""
        data_fine = ""

        inizio_match = re.search(
            r'Data inizio:\s*(\d{1,2})/(\w+)/(\d{4})',
            body,
            re.IGNORECASE
        )
        if inizio_match:
            data_inizio = parse_data_italiana(
                inizio_match.group(1),
                inizio_match.group(2),
                inizio_match.group(3)
            )

        fine_match = re.search(
            r'Data fine:\s*(\d{1,2})/(\w+)/(\d{4})',
            body,
            re.IGNORECASE
        )
        if fine_match:
            data_fine = parse_data_italiana(
                fine_match.group(1),
                fine_match.group(2),
                fine_match.group(3)
            )

        # Estrai tipo fruizione
        fruizione_match = re.search(
            r'Tipo fruizione:\s*(.+?)(?:\n|$)',
            body,
            re.IGNORECASE
        )
        if fruizione_match:
            tipo_fruizione = fruizione_match.group(1).strip()

        # Combina motivo con tipo fruizione se presente
        if tipo_fruizione and not motivo:
            motivo = tipo_fruizione

        return Licenza(
            id=msg_id,
            tipo=tipo_licenza,
            motivo=motivo,
            stato=stato,
            data_inizio=data_inizio,
            data_fine=data_fine,
            email_date=email_date,
            email_id=msg_id
        )

    except Exception as e:
        print(f"  [ERROR] Parsing licenza {msg_id}: {e}")
        return None


def classify_email(subject: str) -> str:
    """Classifica il tipo di email in base all'oggetto."""
    subject_lower = subject.lower()

    if "aggiornamento turno di servizio" in subject_lower:
        return "turno"
    elif "eliminazione turno pianificato" in subject_lower:
        return "eliminazione"
    elif "domanda di licenza" in subject_lower or "domanda di riposo" in subject_lower:
        return "licenza"
    else:
        return "altro"


def load_existing_data() -> Optional[Dict]:
    """Carica i dati esistenti se presenti."""
    if DATA_FILE.exists():
        try:
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return None
    return None


def process_emails(service, anno: int = None, is_first_sync: bool = True) -> Dict[str, Any]:
    """
    Recupera e processa le email di notifica turni.

    - Prima sincronizzazione: tutte le email dall'inizio dell'anno
    - Sincronizzazioni successive: solo ultime 3 settimane

    Gestisce la logica degli aggiornamenti: email più recente vince.
    Usa paginazione per recuperare TUTTE le email (non solo le prime 500).
    """
    if anno is None:
        anno = datetime.now().year

    print(f"Cercando email da {MEMO_SENDER}...")

    # Strutture per raccogliere i dati
    turni_per_data: Dict[str, List[Turno]] = defaultdict(list)
    eliminazioni: List[Dict] = []
    licenze: List[Licenza] = []

    # Costruisci la query in base al tipo di sincronizzazione
    if is_first_sync:
        # Prima sync: tutto l'anno corrente
        query = f'from:{MEMO_SENDER} after:{anno}/01/01'
        print(f"  -> PRIMA SINCRONIZZAZIONE: scarico tutto dal 1 gennaio {anno}")
    else:
        # Sync successiva: solo ultime 3 settimane
        three_weeks_ago = (datetime.now() - timedelta(weeks=3)).strftime('%Y/%m/%d')
        query = f'from:{MEMO_SENDER} after:{three_weeks_ago}'
        print(f"  -> SINCRONIZZAZIONE INCREMENTALE: ultime 3 settimane (dal {three_weeks_ago})")

    try:
        # Usa paginazione per recuperare TUTTE le email
        messages = []
        page_token = None
        page_num = 1

        while True:
            print(f"  Recuperando pagina {page_num}...")

            if page_token:
                results = service.users().messages().list(
                    userId='me',
                    q=query,
                    maxResults=500,
                    pageToken=page_token
                ).execute()
            else:
                results = service.users().messages().list(
                    userId='me',
                    q=query,
                    maxResults=500
                ).execute()

            batch = results.get('messages', [])
            messages.extend(batch)
            print(f"    -> {len(batch)} email in questa pagina (totale: {len(messages)})")

            # Controlla se ci sono altre pagine
            page_token = results.get('nextPageToken')
            if not page_token:
                break
            page_num += 1

        print(f"Trovate {len(messages)} email di notifica totali")

        # Prima passa: raccogli tutte le email
        all_emails = []
        for i, msg_info in enumerate(messages):
            msg_id = msg_info['id']
            print(f"  [{i+1}/{len(messages)}] Recuperando {msg_id}...")

            msg = service.users().messages().get(
                userId='me',
                id=msg_id,
                format='full'
            ).execute()

            body = get_email_body(msg)
            email_date = get_email_date(msg)
            subject = get_email_subject(msg)

            all_emails.append({
                'id': msg_id,
                'body': body,
                'email_date': email_date,
                'subject': subject
            })

        # Ordina per data email (dalla più vecchia alla più recente)
        # Così l'ultima email processata sovrascrive le precedenti
        all_emails.sort(key=lambda x: x['email_date'])

        print(f"\nProcessando email ordinate per data...")

        # Seconda passa: processa in ordine cronologico
        for email in all_emails:
            email_type = classify_email(email['subject'])

            if email_type == "turno":
                turno = parse_turno_servizio(
                    email['body'],
                    email['email_date'],
                    email['id'],
                    email['subject']
                )
                if turno:
                    # Aggiungi o aggiorna il turno per quella data/ora
                    turni_per_data[turno.data].append(turno)
                    print(f"    -> TURNO: {turno.data} {turno.ora_inizio}-{turno.ora_fine} ({turno.durata_ore}h)")

            elif email_type == "eliminazione":
                elim = parse_eliminazione_turno(
                    email['body'],
                    email['email_date'],
                    email['id'],
                    email['subject']
                )
                if elim:
                    eliminazioni.append(elim)
                    print(f"    -> ELIMINAZIONE turno del {elim['data']}")

            elif email_type == "licenza":
                lic = parse_licenza(
                    email['body'],
                    email['email_date'],
                    email['id'],
                    email['subject']
                )
                if lic:
                    licenze.append(lic)
                    print(f"    -> LICENZA {lic.tipo} - {lic.stato}")

    except HttpError as e:
        print(f"Errore API Gmail: {e}")

    return {
        'turni_per_data': turni_per_data,
        'eliminazioni': eliminazioni,
        'licenze': licenze
    }


def turni_si_sovrappongono(t1: Turno, t2: Turno) -> bool:
    """
    Verifica se due turni si sovrappongono (hanno ore in comune).
    Due turni si sovrappongono se uno inizia prima che l'altro finisca.
    Turni consecutivi (uno finisce quando l'altro inizia) NON si sovrappongono.
    """
    # Converti in minuti per facile confronto
    def to_minutes(time_str):
        h, m = map(int, time_str.split(':'))
        return h * 60 + m

    start1 = to_minutes(t1.ora_inizio)
    end1 = to_minutes(t1.ora_fine)
    start2 = to_minutes(t2.ora_inizio)
    end2 = to_minutes(t2.ora_fine)

    # Due turni si sovrappongono se non sono completamente separati
    # Separati = uno finisce prima/quando l'altro inizia
    # t1 finisce prima che t2 inizi, o t2 finisce prima che t1 inizi
    return not (end1 <= start2 or end2 <= start1)


def consolidate_turni(turni_per_data: Dict[str, List[Turno]], eliminazioni: List[Dict]) -> List[Giornata]:
    """
    Consolida i turni per ogni giornata.

    LOGICA CORRETTA:
    - L'ultima email cronologicamente ricevuta è quella effettiva
    - ECCEZIONE: se ci sono turni CONSECUTIVI (es. 8-14 e 14-20), entrambi contano
    - Se due turni si SOVRAPPONGONO, l'ultimo email vince (è un aggiornamento)

    Es: 08:00-14:00 poi 07:00-13:00 → solo 07:00-13:00 (aggiornamento)
    Es: 08:00-14:00 poi 14:00-20:00 → entrambi (turni consecutivi = giornata 8-20)
    """
    giornate = []

    # Raggruppa eliminazioni per data
    eliminazioni_per_data: Dict[str, List[Dict]] = defaultdict(list)
    for elim in eliminazioni:
        eliminazioni_per_data[elim['data']].append(elim)

    for data, turni in sorted(turni_per_data.items()):
        # Ordina turni per data email (dal più vecchio al più recente)
        turni_ordinati = sorted(turni, key=lambda t: t.email_date)

        # Lista turni finali per questa giornata
        turni_consolidati: List[Turno] = []

        for nuovo_turno in turni_ordinati:
            # Controlla se questo turno si sovrappone con turni esistenti
            turni_da_rimuovere = []

            for i, turno_esistente in enumerate(turni_consolidati):
                if turni_si_sovrappongono(nuovo_turno, turno_esistente):
                    # Si sovrappongono: il nuovo (più recente) sostituisce il vecchio
                    turni_da_rimuovere.append(i)

            # Rimuovi turni sovrapposti (in ordine inverso per non sballare gli indici)
            for i in reversed(turni_da_rimuovere):
                rimosso = turni_consolidati.pop(i)
                rimosso.stato = 'eliminato'
                print(f"  [SOVRAPP] {data}: {rimosso.ora_inizio}-{rimosso.ora_fine} sostituito da {nuovo_turno.ora_inizio}-{nuovo_turno.ora_fine}")

            # Aggiungi il nuovo turno
            turni_consolidati.append(nuovo_turno)

        # Applica eliminazioni per questa data
        if data in eliminazioni_per_data:
            for elim in eliminazioni_per_data[data]:
                ora_elim = elim.get('ora_inizio', '')
                ora_fine_elim = elim.get('ora_fine', '')

                for turno in turni_consolidati:
                    if turno.stato == 'eliminato':
                        continue

                    # Confronta se l'eliminazione corrisponde a questo turno
                    if ora_elim:
                        # Eliminazione specifica: deve corrispondere l'ora inizio
                        if turno.ora_inizio == ora_elim:
                            if elim['email_date'] > turno.email_date:
                                turno.stato = 'eliminato'
                                print(f"  [ELIM] Turno {data} {ora_elim} eliminato")
                    else:
                        # Eliminazione generica: elimina tutti i turni più vecchi
                        if elim['email_date'] > turno.email_date:
                            turno.stato = 'eliminato'
                            print(f"  [ELIM] Turno {data} {turno.ora_inizio} eliminato (generico)")

        # Separa turni attivi
        turni_attivi = [t for t in turni_consolidati if t.stato == 'attivo']

        # IMPORTANTE: Le ore di ASSENZA non vanno conteggiate!
        turni_lavorativi = [t for t in turni_attivi if t.tipo != 'ASSENZA']
        ore_totali = sum(t.durata_ore for t in turni_lavorativi)

        # Ricalcola straordinario considerando il totale giornaliero
        ore_ordinarie = min(ore_totali, ORE_TURNO_STANDARD)
        ore_straordinario = max(0, ore_totali - ORE_TURNO_STANDARD)

        giornata = Giornata(
            data=data,
            turni=turni_consolidati,  # Include anche eliminati per tracciabilità
            ore_totali=round(ore_totali, 2),
            ore_ordinarie=round(ore_ordinarie, 2),
            ore_straordinario=round(ore_straordinario, 2)
        )
        giornate.append(giornata)

    # Ordina per data (più recente prima)
    giornate.sort(key=lambda g: g.data, reverse=True)

    return giornate


def add_missing_rest_days(giornate: List[Giornata], anno: int = 2025) -> List[Giornata]:
    """
    Aggiunge automaticamente i giorni di riposo mancanti:
    - Domeniche senza servizio → Riposo Settimanale (RS)
    - Festività senza servizio → Riposo Festivo (RF)

    Logica:
    - Ogni settimana ha diritto a 1 giorno di riposo settimanale
    - Le festività nazionali sono riposi festivi se non lavorati
    """
    print("\nAggiunta giorni di riposo mancanti...")

    # Crea un set delle date già presenti
    date_esistenti = {g.data for g in giornate}

    # Calcola la data di oggi per non aggiungere giorni futuri
    oggi = datetime.now().strftime('%Y-%m-%d')

    # Ottieni festività e domeniche
    festivi = get_festivi_italiani(anno)
    domeniche = get_all_sundays(anno)

    giorni_aggiunti = 0

    # 1. Aggiungi festività mancanti come "Riposo Festivo"
    for data, nome_festivo in festivi.items():
        if data not in date_esistenti and data <= oggi and data >= f"{anno}-01-01":
            # Crea turno fittizio per il riposo festivo
            turno = Turno(
                id=f"{data}_RF_auto",
                tipo="ASSENZA",
                dettaglio=f"Riposo festivita' ({nome_festivo})",
                matricola=EMPLOYEE_ID,
                data=data,
                ora_inizio="00:00",
                ora_fine="23:59",
                durata_ore=0,
                is_straordinario=False,
                ore_ordinarie=0,
                ore_straordinario=0,
                email_date=datetime.now().isoformat(),
                email_id="auto_generated",
                stato="attivo"
            )

            giornata = Giornata(
                data=data,
                turni=[turno],
                ore_totali=0,
                ore_ordinarie=0,
                ore_straordinario=0
            )
            giornate.append(giornata)
            date_esistenti.add(data)
            giorni_aggiunti += 1
            print(f"  + {data}: Riposo Festivo ({nome_festivo})")

    # 2. Aggiungi domeniche mancanti come "Riposo Settimanale"
    for data in domeniche:
        if data not in date_esistenti and data <= oggi and data >= f"{anno}-01-01":
            # Controlla se non è già una festività (già aggiunta sopra)
            if data not in festivi:
                turno = Turno(
                    id=f"{data}_RS_auto",
                    tipo="ASSENZA",
                    dettaglio="Riposo settimanale",
                    matricola=EMPLOYEE_ID,
                    data=data,
                    ora_inizio="00:00",
                    ora_fine="23:59",
                    durata_ore=0,
                    is_straordinario=False,
                    ore_ordinarie=0,
                    ore_straordinario=0,
                    email_date=datetime.now().isoformat(),
                    email_id="auto_generated",
                    stato="attivo"
                )

                giornata = Giornata(
                    data=data,
                    turni=[turno],
                    ore_totali=0,
                    ore_ordinarie=0,
                    ore_straordinario=0
                )
                giornate.append(giornata)
                date_esistenti.add(data)
                giorni_aggiunti += 1
                print(f"  + {data}: Riposo Settimanale")

    # Riordina per data
    giornate.sort(key=lambda g: g.data, reverse=True)

    print(f"  -> Aggiunti {giorni_aggiunti} giorni di riposo")

    return giornate


def expand_licenses_to_giornate(giornate: List[Giornata], licenze: List[Licenza]) -> List[Giornata]:
    """
    Espande le licenze approvate in giornate.
    Per ogni licenza approvata, crea una giornata per ogni giorno del periodo.
    """
    print("\nEspansione licenze approvate in giornate...")

    # Set delle date già esistenti
    date_esistenti = set(g.data for g in giornate)

    # Filtra solo licenze approvate (stato finale)
    licenze_approvate = [l for l in licenze if l.stato == 'Approvata']

    # Rimuovi duplicati (stessa licenza può avere più record per stati diversi)
    licenze_uniche = {}
    for lic in licenze_approvate:
        key = f"{lic.tipo}_{lic.data_inizio}_{lic.data_fine}"
        licenze_uniche[key] = lic

    giorni_aggiunti = 0

    for lic in licenze_uniche.values():
        # Parse date
        try:
            data_inizio = datetime.strptime(lic.data_inizio, '%Y-%m-%d')
            data_fine = datetime.strptime(lic.data_fine, '%Y-%m-%d')
        except:
            print(f"  [SKIP] Date non valide: {lic.data_inizio} - {lic.data_fine}")
            continue

        # Genera ogni giorno nel range
        current = data_inizio
        while current <= data_fine:
            data_str = current.strftime('%Y-%m-%d')

            # Solo se non esiste già una giornata
            if data_str not in date_esistenti:
                # Determina il dettaglio in base al tipo di licenza
                tipo_licenza = lic.tipo or 'ordinaria'
                if tipo_licenza == 'ordinaria':
                    dettaglio = 'Licenza ordinaria'
                elif tipo_licenza == 'straordinaria':
                    dettaglio = 'Licenza straordinaria'
                elif tipo_licenza == 'speciale':
                    dettaglio = 'Licenza speciale'
                elif tipo_licenza == 'riposo_donatori':
                    dettaglio = 'Riposo donatore sangue'
                else:
                    dettaglio = f'Licenza {tipo_licenza}'

                turno = Turno(
                    id=f"lic_{data_str}_{tipo_licenza}",
                    tipo='ASSENZA',
                    dettaglio=dettaglio,
                    matricola=EMPLOYEE_ID,
                    data=data_str,
                    ora_inizio='00:00',
                    ora_fine='23:59',
                    durata_ore=0,
                    is_straordinario=False,
                    ore_ordinarie=0,
                    ore_straordinario=0,
                    email_date=datetime.now().isoformat(),
                    email_id=f'licenza_{tipo_licenza}',
                    stato='attivo'
                )

                giornata = Giornata(
                    data=data_str,
                    turni=[turno],
                    ore_totali=0,
                    ore_ordinarie=0,
                    ore_straordinario=0,
                    is_licenza=True,
                    tipo_licenza=tipo_licenza
                )
                giornate.append(giornata)
                date_esistenti.add(data_str)
                giorni_aggiunti += 1

            current += timedelta(days=1)

    # Riordina per data
    giornate.sort(key=lambda g: g.data, reverse=True)

    print(f"  -> Aggiunti {giorni_aggiunti} giorni di licenza")

    return giornate


def get_festivita_italiane(anno: int) -> set:
    """
    Restituisce le festività italiane per l'anno specificato.
    Include festività fisse + Pasqua e Lunedì dell'Angelo.
    """
    # Festività fisse
    festivita = {
        f"{anno}-01-01",  # Capodanno
        f"{anno}-01-06",  # Epifania
        f"{anno}-04-25",  # Liberazione
        f"{anno}-05-01",  # Festa del Lavoro
        f"{anno}-06-02",  # Festa della Repubblica
        f"{anno}-08-15",  # Ferragosto
        f"{anno}-11-01",  # Ognissanti
        f"{anno}-12-08",  # Immacolata
        f"{anno}-12-25",  # Natale
        f"{anno}-12-26",  # Santo Stefano
    }

    # Calcolo Pasqua (algoritmo di Gauss)
    a = anno % 19
    b = anno // 100
    c = anno % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mese = (h + l - 7 * m + 114) // 31
    giorno = ((h + l - 7 * m + 114) % 31) + 1

    pasqua = date(anno, mese, giorno)
    lunedi_angelo = pasqua + timedelta(days=1)

    festivita.add(pasqua.strftime("%Y-%m-%d"))
    festivita.add(lunedi_angelo.strftime("%Y-%m-%d"))

    return festivita


def is_giorno_festivo(data_str: str) -> bool:
    """
    Verifica se una data è un giorno festivo (domenica o festività nazionale).
    """
    try:
        data = datetime.strptime(data_str, "%Y-%m-%d").date()
        anno = data.year

        # Domenica = 6
        if data.weekday() == 6:
            return True

        # Festività nazionali
        festivita = get_festivita_italiane(anno)
        return data_str in festivita

    except:
        return False


def calcola_ore_per_fascia(ora_inizio: str, ora_fine: str, ore_straordinario: float, is_festivo: bool) -> dict:
    """
    Calcola le ore di straordinario suddivise per fascia oraria.

    Fasce:
    - Diurno: 06:00 - 22:00
    - Notturno: 22:00 - 06:00

    Ritorna dict con:
    - diurno: ore feriali diurne
    - notturno: ore feriali notturne
    - festivo_diurno: ore festive diurne
    - festivo_notturno: ore festive notturne
    """
    result = {
        'diurno': 0,
        'notturno': 0,
        'festivo_diurno': 0,
        'festivo_notturno': 0
    }

    if ore_straordinario <= 0:
        return result

    try:
        def to_minutes(time_str):
            h, m = map(int, time_str.split(':'))
            return h * 60 + m

        start = to_minutes(ora_inizio)
        end = to_minutes(ora_fine)

        # Se il turno attraversa la mezzanotte
        if end <= start:
            end += 24 * 60

        # Limiti fasce in minuti
        DIURNO_START = 6 * 60   # 06:00
        DIURNO_END = 22 * 60    # 22:00

        # Calcola minuti diurni e notturni del turno totale
        minuti_diurni = 0
        minuti_notturni = 0

        for minuto in range(start, end):
            minuto_normalizzato = minuto % (24 * 60)
            if DIURNO_START <= minuto_normalizzato < DIURNO_END:
                minuti_diurni += 1
            else:
                minuti_notturni += 1

        # Calcola la proporzione di ore diurne/notturne
        totale_minuti = minuti_diurni + minuti_notturni
        if totale_minuti > 0:
            prop_diurno = minuti_diurni / totale_minuti
            prop_notturno = minuti_notturni / totale_minuti

            ore_diurne = ore_straordinario * prop_diurno
            ore_notturne = ore_straordinario * prop_notturno

            if is_festivo:
                result['festivo_diurno'] = round(ore_diurne, 2)
                result['festivo_notturno'] = round(ore_notturne, 2)
            else:
                result['diurno'] = round(ore_diurne, 2)
                result['notturno'] = round(ore_notturne, 2)

    except Exception as e:
        # In caso di errore, assegna tutto come diurno feriale/festivo
        if is_festivo:
            result['festivo_diurno'] = ore_straordinario
        else:
            result['diurno'] = ore_straordinario

    return result


def calculate_stats(giornate: List[Giornata], licenze: List[Licenza]) -> Dict:
    """Calcola statistiche complete."""
    if not giornate:
        return {
            'ore_totali': 0,
            'ore_ordinarie': 0,
            'ore_straordinario': 0,
            'giorni_lavorati': 0,
            'media_ore_giorno': 0,
            'per_tipo': {},
            'per_mese': {},
            'licenze_per_tipo': {}
        }

    # Statistiche base
    ore_totali = sum(g.ore_totali for g in giornate)
    ore_ordinarie = sum(g.ore_ordinarie for g in giornate)
    ore_straordinario = sum(g.ore_straordinario for g in giornate)
    giorni_lavorati = len([g for g in giornate if g.ore_totali > 0])

    # Per tipo di servizio (ASSENZA tracciata separatamente, senza ore)
    per_tipo = defaultdict(lambda: {'count': 0, 'ore': 0})
    assenze_count = 0
    for g in giornate:
        for t in g.turni:
            if t.stato == 'attivo':
                per_tipo[t.tipo]['count'] += 1
                # Le ore di ASSENZA non vanno conteggiate
                if t.tipo != 'ASSENZA':
                    per_tipo[t.tipo]['ore'] += t.durata_ore
                else:
                    assenze_count += 1

    # Per mese (incluse turnazioni esterne e breakdown straordinario)
    per_mese = defaultdict(lambda: {
        'giorni': 0, 'ore': 0, 'ore_ordinarie': 0, 'ore_straordinario': 0,
        'turnazioni_esterne': 0,
        'straord_diurno': 0, 'straord_notturno': 0,
        'straord_festivo_diurno': 0, 'straord_festivo_notturno': 0,
        'recuperi_mese': 0,  # Recupero ore prestate nel mese in corso
        'recuperi_non_retribuiti': 0  # Recupero ore non retribuite
    })
    turnazioni_esterne_totali = 0

    # Totali recuperi
    totale_recuperi_mese = 0  # -6h per ogni recupero dal monte ore del mese
    totale_recuperi_non_retribuiti = 0  # -6h per ogni recupero dal monte ore accumulato

    # Totali breakdown straordinario
    totale_straord_diurno = 0
    totale_straord_notturno = 0
    totale_straord_festivo_diurno = 0
    totale_straord_festivo_notturno = 0

    for g in giornate:
        mese = g.data[:7]

        # Conta recuperi (anche per giornate senza ore lavorate)
        for t in g.turni:
            if t.stato == 'attivo' and t.tipo == 'ASSENZA':
                dettaglio_lower = t.dettaglio.lower()
                if 'recupero di ore prestate nel mese' in dettaglio_lower:
                    # Recupero ore del mese in corso: -6h dal monte ore mensile
                    per_mese[mese]['recuperi_mese'] += 1
                    totale_recuperi_mese += 1
                elif 'recupero di ore non retribuite' in dettaglio_lower:
                    # Recupero ore non retribuite: -6h dal monte ore accumulato
                    per_mese[mese]['recuperi_non_retribuiti'] += 1
                    totale_recuperi_non_retribuiti += 1

        if g.ore_totali > 0:
            per_mese[mese]['giorni'] += 1
            per_mese[mese]['ore'] += g.ore_totali
            per_mese[mese]['ore_ordinarie'] += g.ore_ordinarie
            per_mese[mese]['ore_straordinario'] += g.ore_straordinario

            # Verifica se è giorno festivo
            is_festivo = is_giorno_festivo(g.data)

            # Conta turnazioni esterne e calcola breakdown straordinario
            turni_attivi = [t for t in g.turni if t.stato == 'attivo']
            # Considera come turno esterno se ha dettaglio specifico o è un servizio di presenza
            has_turno_esterno = any(
                'scorta' in t.dettaglio.lower() or
                'esterno' in t.dettaglio.lower() or
                'accompagn' in t.dettaglio.lower()
                for t in turni_attivi
            )

            if has_turno_esterno:
                if g.ore_totali > 12:
                    per_mese[mese]['turnazioni_esterne'] += 2
                    turnazioni_esterne_totali += 2
                else:
                    per_mese[mese]['turnazioni_esterne'] += 1
                    turnazioni_esterne_totali += 1

            # Calcola breakdown straordinario per la giornata
            # Lo straordinario sono le ore OLTRE le prime 6 ordinarie
            # Quindi sono le ULTIME ore della giornata lavorativa
            if g.ore_straordinario > 0:
                # Raccogli tutti i minuti lavorati ordinati cronologicamente
                minuti_lavorati = []

                def to_minutes(time_str):
                    h, m = map(int, time_str.split(':'))
                    return h * 60 + m

                DIURNO_START = 6 * 60   # 06:00
                DIURNO_END = 22 * 60    # 22:00

                for t in turni_attivi:
                    try:
                        start = to_minutes(t.ora_inizio)
                        end = to_minutes(t.ora_fine)
                        if end <= start:
                            end += 24 * 60

                        for minuto in range(start, end):
                            minuti_lavorati.append(minuto)
                    except:
                        pass

                # Ordina i minuti cronologicamente
                minuti_lavorati.sort()

                # Le prime 360 minuti (6h) sono ordinarie
                # I minuti oltre sono straordinario
                minuti_straordinario = minuti_lavorati[360:] if len(minuti_lavorati) > 360 else []

                # Conta minuti straordinario diurni e notturni
                straord_min_diurno = 0
                straord_min_notturno = 0

                for minuto in minuti_straordinario:
                    minuto_norm = minuto % (24 * 60)
                    if DIURNO_START <= minuto_norm < DIURNO_END:
                        straord_min_diurno += 1
                    else:
                        straord_min_notturno += 1

                # Converti in ore e arrotonda a 0.5h
                straord_diurno = round(straord_min_diurno / 60 * 2) / 2
                straord_notturno = round(straord_min_notturno / 60 * 2) / 2

                if is_festivo:
                    per_mese[mese]['straord_festivo_diurno'] += straord_diurno
                    per_mese[mese]['straord_festivo_notturno'] += straord_notturno
                    totale_straord_festivo_diurno += straord_diurno
                    totale_straord_festivo_notturno += straord_notturno
                else:
                    per_mese[mese]['straord_diurno'] += straord_diurno
                    per_mese[mese]['straord_notturno'] += straord_notturno
                    totale_straord_diurno += straord_diurno
                    totale_straord_notturno += straord_notturno

    # Licenze per tipo e stato
    licenze_per_tipo = defaultdict(lambda: defaultdict(int))
    for lic in licenze:
        licenze_per_tipo[lic.tipo][lic.stato] += 1

    return {
        'ore_totali': round(ore_totali, 2),
        'ore_ordinarie': round(ore_ordinarie, 2),
        'ore_straordinario': round(ore_straordinario, 2),
        'giorni_lavorati': giorni_lavorati,
        'media_ore_giorno': round(ore_totali / giorni_lavorati, 2) if giorni_lavorati > 0 else 0,
        'turnazioni_esterne': turnazioni_esterne_totali,
        'straord_diurno': round(totale_straord_diurno, 2),
        'straord_notturno': round(totale_straord_notturno, 2),
        'straord_festivo_diurno': round(totale_straord_festivo_diurno, 2),
        'straord_festivo_notturno': round(totale_straord_festivo_notturno, 2),
        'recuperi_mese_totale': totale_recuperi_mese,  # N. giorni di recupero ore mese
        'recuperi_non_retribuiti_totale': totale_recuperi_non_retribuiti,  # N. giorni di recupero ore non retribuite
        'ore_recuperate_mese': totale_recuperi_mese * 6,  # Ore scalate dallo straord mensile
        'ore_recuperate_non_retribuite': totale_recuperi_non_retribuiti * 6,  # Ore scalate dal monte ore
        'per_tipo': dict(per_tipo),
        'per_mese': dict(sorted(per_mese.items(), reverse=True)),
        'licenze_per_tipo': {k: dict(v) for k, v in licenze_per_tipo.items()}
    }


def save_data(giornate: List[Giornata], licenze: List[Licenza], stats: Dict, anno: int = None):
    """Salva tutti i dati nel file JSON."""
    if anno is None:
        anno = datetime.now().year

    # Lista archivi disponibili
    archives = []
    data_dir = BASE_DIR / 'data'
    for f in data_dir.glob('archivio_*.json'):
        try:
            arch_anno = int(f.stem.replace('archivio_', ''))
            archives.append(arch_anno)
        except:
            pass

    output = {
        'anno': anno,
        'last_update': datetime.now().isoformat(),
        'total_giorni': len(giornate),
        'total_licenze': len(licenze),
        'archives': sorted(archives),
        'stats': stats,
        'giornate': [g.to_dict() for g in giornate],
        'licenze': [l.to_dict() for l in licenze],
        # Mantieni anche il formato "servizi" per compatibilità con la dashboard
        'servizi': []
    }

    # Converti giornate in formato servizi per la dashboard esistente
    for g in giornate:
        for t in g.turni:
            if t.stato == 'attivo':
                output['servizi'].append({
                    'id': t.id,
                    'tipo_servizio': t.tipo,
                    'dettaglio': sanitize_dettaglio(t.dettaglio),
                    'data_inizio': t.data,
                    'ora_inizio': t.ora_inizio,
                    'data_fine': t.data,
                    'ora_fine': t.ora_fine,
                    'durata_ore': t.durata_ore,
                    'is_straordinario': t.is_straordinario,
                    'ore_ordinarie': t.ore_ordinarie,
                    'ore_straordinario': t.ore_straordinario,
                    'email_date': t.email_date
                })

    output['total_servizi'] = len(output['servizi'])

    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    # Copia anche gli archivi nella cartella docs per la dashboard
    docs_dir = BASE_DIR / 'docs'
    for arch_file in data_dir.glob('archivio_*.json'):
        import shutil
        dest = docs_dir / arch_file.name
        shutil.copy(arch_file, dest)

    print(f"\nDati salvati in {DATA_FILE}")
    if archives:
        print(f"Archivi disponibili: {archives}")
    return output


def merge_with_existing(existing_data: Dict, new_giornate: List[Giornata], new_licenze: List[Licenza]) -> tuple:
    """
    Unisce i dati esistenti con i nuovi.
    Per le ultime 3 settimane, i nuovi dati sovrascrivono.
    Per i dati più vecchi, mantiene quelli esistenti.
    """
    # Crea un dizionario delle giornate esistenti
    existing_giornate_dict = {}
    if 'giornate' in existing_data:
        for g_data in existing_data['giornate']:
            existing_giornate_dict[g_data['data']] = g_data

    # Crea un dizionario delle licenze esistenti
    existing_licenze_dict = {}
    if 'licenze' in existing_data:
        for l_data in existing_data['licenze']:
            key = f"{l_data['id']}_{l_data['stato']}"
            existing_licenze_dict[key] = l_data

    # Calcola la data limite (3 settimane fa)
    three_weeks_ago = (datetime.now() - timedelta(weeks=3)).strftime('%Y-%m-%d')

    # Aggiorna con i nuovi dati (ultime 3 settimane)
    for g in new_giornate:
        existing_giornate_dict[g.data] = g.to_dict()

    # Aggiorna licenze
    for lic in new_licenze:
        key = f"{lic.id}_{lic.stato}"
        existing_licenze_dict[key] = lic.to_dict()

    # Riconverti in liste
    merged_giornate_data = list(existing_giornate_dict.values())
    merged_licenze_data = list(existing_licenze_dict.values())

    # Ordina giornate per data (più recente prima)
    merged_giornate_data.sort(key=lambda x: x['data'], reverse=True)

    # Riconverti in oggetti Giornata
    merged_giornate = []
    for g_data in merged_giornate_data:
        turni = []
        for t_data in g_data.get('turni', []):
            # Aggiungi matricola di default se mancante (rimossa per privacy)
            if 'matricola' not in t_data:
                t_data['matricola'] = EMPLOYEE_ID
            turni.append(Turno(**t_data))
        giornata = Giornata(
            data=g_data['data'],
            turni=turni,
            ore_totali=g_data.get('ore_totali', 0),
            ore_ordinarie=g_data.get('ore_ordinarie', 0),
            ore_straordinario=g_data.get('ore_straordinario', 0),
            is_licenza=g_data.get('is_licenza', False),
            tipo_licenza=g_data.get('tipo_licenza', '')
        )
        merged_giornate.append(giornata)

    # Riconverti licenze
    merged_licenze = []
    for l_data in merged_licenze_data:
        merged_licenze.append(Licenza(**l_data))

    print(f"  -> Dati uniti: {len(merged_giornate)} giornate, {len(merged_licenze)} licenze")

    return merged_giornate, merged_licenze


def get_archive_file(anno: int) -> Path:
    """Restituisce il path del file archivio per un anno specifico."""
    return BASE_DIR / 'data' / f'archivio_{anno}.json'


def archive_year(anno: int):
    """
    Archivia i dati di un anno specifico.
    Copia i dati dell'anno nel file archivio e li rimuove dal file principale.
    """
    existing_data = load_existing_data()
    if not existing_data:
        print(f"[WARN] Nessun dato da archiviare")
        return

    # Filtra giornate dell'anno da archiviare
    anno_str = str(anno)
    giornate_anno = [g for g in existing_data.get('giornate', [])
                     if g['data'].startswith(anno_str)]
    licenze_anno = [l for l in existing_data.get('licenze', [])
                    if l.get('data_inizio', '').startswith(anno_str)]

    if not giornate_anno:
        print(f"[WARN] Nessuna giornata trovata per l'anno {anno}")
        return

    # Calcola statistiche per l'anno
    giornate_obj = []
    for g_data in giornate_anno:
        turni_data = g_data.get('turni', [])
        for t in turni_data:
            if 'matricola' not in t:
                t['matricola'] = EMPLOYEE_ID
        turni = [Turno(**t) for t in turni_data]
        giornate_obj.append(Giornata(
            data=g_data['data'],
            turni=turni,
            ore_totali=g_data.get('ore_totali', 0),
            ore_ordinarie=g_data.get('ore_ordinarie', 0),
            ore_straordinario=g_data.get('ore_straordinario', 0)
        ))
    licenze_obj = [Licenza(**l) for l in licenze_anno]
    stats_anno = calculate_stats(giornate_obj, licenze_obj)

    # Salva archivio
    archivio = {
        'anno': anno,
        'archived_at': datetime.now().isoformat(),
        'total_giorni': len(giornate_anno),
        'total_licenze': len(licenze_anno),
        'stats': stats_anno,
        'giornate': giornate_anno,
        'licenze': licenze_anno
    }

    archive_file = get_archive_file(anno)
    with open(archive_file, 'w', encoding='utf-8') as f:
        json.dump(archivio, f, ensure_ascii=False, indent=2)

    print(f"[OK] Anno {anno} archiviato in {archive_file}")
    print(f"     {len(giornate_anno)} giornate, {len(licenze_anno)} licenze")

    return archivio


def load_archive(anno: int) -> Optional[Dict]:
    """Carica l'archivio di un anno specifico."""
    archive_file = get_archive_file(anno)
    if archive_file.exists():
        with open(archive_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None


def list_archives() -> List[int]:
    """Restituisce la lista degli anni archiviati."""
    data_dir = BASE_DIR / 'data'
    archives = []
    for f in data_dir.glob('archivio_*.json'):
        try:
            anno = int(f.stem.replace('archivio_', ''))
            archives.append(anno)
        except:
            pass
    return sorted(archives)


def main():
    """Funzione principale."""
    anno_corrente = datetime.now().year

    print("=" * 60)
    print(f"Shift Email Fetcher - v3.0 (Anno {anno_corrente})")
    print("=" * 60)

    # Controlla se è la prima sincronizzazione
    existing_data = load_existing_data()

    # Verifica se ci sono dati di anni precedenti da archiviare
    if existing_data:
        anni_presenti = set()
        for g in existing_data.get('giornate', []):
            try:
                anno = int(g['data'][:4])
                anni_presenti.add(anno)
            except:
                pass

        # Archivia anni precedenti
        for anno in anni_presenti:
            if anno < anno_corrente:
                archive_file = get_archive_file(anno)
                if not archive_file.exists():
                    print(f"\n[AUTO-ARCHIVE] Archiviando anno {anno}...")
                    archive_year(anno)

    # Determina se è la prima sync per l'anno corrente
    is_first_sync = existing_data is None or not any(
        g['data'].startswith(str(anno_corrente))
        for g in existing_data.get('giornate', [])
    )

    if is_first_sync:
        print(f"\n[INFO] Prima sincronizzazione {anno_corrente} - scarico tutte le email")
    else:
        giornate_anno = [g for g in existing_data.get('giornate', [])
                         if g['data'].startswith(str(anno_corrente))]
        print(f"\n[INFO] Sincronizzazione incrementale - ultime 3 settimane")
        print(f"       Dati esistenti {anno_corrente}: {len(giornate_anno)} giornate")

    # Connetti a Gmail
    print("\nConnessione a Gmail...")
    service = get_gmail_service()
    print("Connesso!")

    # Processa email (solo anno corrente)
    data = process_emails(service, anno=anno_corrente, is_first_sync=is_first_sync)

    # Consolida turni
    print("\nConsolidamento turni...")
    giornate = consolidate_turni(data['turni_per_data'], data['eliminazioni'])
    print(f"Giornate nuove/aggiornate: {len(giornate)}")

    # Se non è la prima sync, unisci con i dati esistenti (solo anno corrente)
    licenze = data['licenze']
    if not is_first_sync and existing_data:
        # Filtra solo dati dell'anno corrente
        existing_data_filtered = {
            'giornate': [g for g in existing_data.get('giornate', [])
                         if g['data'].startswith(str(anno_corrente))],
            'licenze': [l for l in existing_data.get('licenze', [])
                        if l.get('data_inizio', '').startswith(str(anno_corrente))]
        }
        print("\nUnione con dati esistenti...")
        giornate, licenze = merge_with_existing(existing_data_filtered, giornate, licenze)

    # Espandi licenze approvate in giornate
    giornate = expand_licenses_to_giornate(giornate, licenze)

    # Calcola statistiche
    stats = calculate_stats(giornate, licenze)

    # Salva (solo anno corrente nel file principale)
    output = save_data(giornate, licenze, stats)

    # Aggiungi info sugli archivi disponibili
    archives = list_archives()
    if archives:
        print(f"\n[INFO] Archivi disponibili: {archives}")

    # Riepilogo
    print("\n" + "=" * 60)
    print("RIEPILOGO")
    print("=" * 60)
    print(f"Giorni lavorati: {stats['giorni_lavorati']}")
    print(f"Ore totali: {stats['ore_totali']}h")
    print(f"  - Ordinarie: {stats['ore_ordinarie']}h")
    print(f"  - Straordinario: {stats['ore_straordinario']}h")
    print(f"Media ore/giorno: {stats['media_ore_giorno']}h")

    print("\nPer tipo servizio:")
    for tipo, data in stats.get('per_tipo', {}).items():
        print(f"  - {tipo}: {data['count']} turni, {data['ore']:.1f}h")

    print("\nPer mese:")
    for mese, data in list(stats.get('per_mese', {}).items())[:3]:
        print(f"  - {mese}: {data['giorni']} giorni, {data['ore']:.1f}h totali ({data['ore_straordinario']:.1f}h straord.)")

    if stats.get('licenze_per_tipo'):
        print("\nLicenze:")
        for tipo, stati in stats['licenze_per_tipo'].items():
            print(f"  - {tipo}: {stati}")


if __name__ == '__main__':
    main()
