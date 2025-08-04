import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from ..models import Livraison
import pytz

def import_livraisons_from_sheet2():
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        'importer/credentials/sheetimporterproject-8a59547703c3.json', scope)
    client = gspread.authorize(creds)

    sheet = client.open_by_key("1jl2rRck8m4TeSaV91JjJ89kbIxXfNIWshQf_lAwbRYw").sheet1
    rows = sheet.get_all_records()
    for sheet in client.openall():
        print(sheet.title)
    tz = pytz.timezone("Africa/Casablanca")  # Adjust as needed
    for row in rows:
        timestamp_str = row.get("timestamp") or row.get("Timestamp")  # handle both cases
        if not timestamp_str:
            continue  # Skip if timestamp missing

        try:
            # Convert string "14/07/2025 13:24:11" → aware datetime
            naive_dt = datetime.strptime(timestamp_str, "%d/%m/%Y %H:%M:%S")
            aware_dt = tz.localize(naive_dt)
        except Exception as e:
            print(f"Skipping row due to timestamp format issue: {e}")
            continue

        if Livraison.objects.filter(timestamp=aware_dt).exists():
            continue  # Skip duplicates

        Livraison.objects.create(
            timestamp=aware_dt,
            nom_chauffeur=row.get("nom_chauffeur") or row.get("NOM DE CHAUFFEUR"),
            client=row.get("clients") or row.get("CLIENTS"),
            chargements=row.get("LIEU DE CHARGE") or row.get("LIEU DE CHARGE"),
            date_chargement=parse_date(row.get("date_chargement") or row.get("DATE DE CHARGE")),
            dechargement=row.get("lieu de charge") or row.get("LIEU DE DECHARGE"),
            bon_livraison=row.get("bon_livraison") or row.get("BON DE LIVRAISON"),
            tarif=row.get("PRIX DE VOYAGE") or row.get("PRIX DE VOYAGE"),
            deplacement=parse_decimal(row.get("deplacement") or row.get("DEPLACEMENT")),
            avance=parse_decimal(row.get("avance")),
            charge_variable=parse_decimal(row.get("CHARGE VARIABLE")),
            prix_cv=parse_decimal(row.get("prix de ch.v") or row.get("PRIX DE CH.V")),
            operateur=row.get("operateur") or row.get("OPERATEUR"),
            commercial=row.get("commercial") or row.get("COMMERCIAL"),
            ice=row.get("ice") or row.get("ICE"),
            qte=parse_int(row.get("qte") or row.get("QTE")),
            nom_destinataire=row.get("nom_destinataire") or row.get("NOM DE DESTINATAIRE"),
            facturation=row.get("facturation") or row.get("Facturation")
        )

    print("✅ Livraisons imported successfully!")

def parse_date(date_str):
    try:
        return datetime.strptime(date_str, "%d/%m/%Y").date()
    except:
        return None

def parse_datetime(dt_str):
    if not dt_str:
        return None
    try:
        return datetime.strptime(dt_str, '%d/%m/%Y %H:%M:%S')
    except Exception:
        return None

def parse_decimal(value):
    try:
        return float(str(value).replace(",", "."))
    except:
        return 0.0

def parse_int(value):
    try:
        return int(float(value))
    except:
        return 0
