import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from ..models import Livraison
import pytz

def import_livraisons_from_sheet():
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        'importer/credentials/sheetimporterproject-468211-d4bfa68f711f.json', scope)
    client = gspread.authorize(creds)

    sheet = client.open_by_key("1aAVGSHrjGE82Py3MiL5MUtS_jo2JwKp-O7lR5-VrVGE").sheet1
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
            client=row.get("client") or row.get("CLIENTS"),
            chargements=row.get("chargements") or row.get("CHARGEMENTS"),
            date_chargement=parse_date_flexible(row.get("date_chargement") or row.get("DATE DE CHAR")),
            dechargement=row.get("dechargement") or row.get("DECHARGEMENT"),
            bon_livraison=row.get("bon_livraison") or row.get("bon de livraison"),
            tarif=row.get("tarif") or row.get("Tarif"),
            deplacement=parse_decimal(row.get("deplacement") or row.get("DEPLACEMENT")),
            avance=parse_decimal(row.get("avance")),
            charge_variable=parse_decimal(row.get("charge_variable")),
            prix_cv=parse_decimal(row.get("prix_cv") or row.get("prix de C.V")),
            operateur=row.get("operateur") or row.get("OPERATEUR"),
            commercial=row.get("commercial") or row.get("COMMERCIAL"),
            ice=row.get("ice") or row.get("ICE"),
            qte=parse_int(row.get("qte") or row.get("QTE")),
            nom_destinataire=row.get("nom_destinataire") or row.get("NOM DE DESTINATAIRE"),
            facturation=row.get("facturation") or row.get("Facturation")
        )

    print("✅ Livraisons imported successfully!")


def parse_date_flexible(date_str):
    """Parse date with multiple format support - similar to timestamp parsing logic"""
    if not date_str:
        return None
    
    # Remove extra whitespace
    date_str = date_str.strip()
    if not date_str:
        return None
    
    # Try multiple date formats - both American and European
    formats_to_try = [
        "%m/%d/%Y",      # 6/30/2025 (American format)
        "%d/%m/%Y",      # 18/07/2025 (European format)
        "%Y-%m-%d",      # 2025-07-18 (ISO format)
        "%d-%m-%Y",      # 18-07-2025 (European with dashes)
        "%m-%d-%Y",      # 07-18-2025 (American with dashes)
        "%d.%m.%Y",      # 18.07.2025 (European with dots)
        "%m.%d.%Y",      # 07.18.2025 (American with dots)
        "%Y/%m/%d",      # 2025/07/18 (ISO with slashes)
    ]
    
    for fmt in formats_to_try:
        try:
            parsed_date = datetime.strptime(date_str, fmt).date()
            print(f"   📅 Date '{date_str}' parsed with format '{fmt}': {parsed_date}")
            return parsed_date
        except ValueError:
            continue
    
    print(f"   ❌ Could not parse date: '{date_str}' with any format")
    return None

def parse_datetime(dt_str):
    if not dt_str:
        return None
    try:
        return datetime.strptime(dt_str, '%d/%m/%Y %H:%M:%S')
    except Exception:
        return None

def import_livraisons_from_sheet2():        
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        'importer/credentials/sheetimporterproject-468211-d4bfa68f711f.json', scope)
    client = gspread.authorize(creds)

    sheet = client.open_by_key("1oxQxHFF3jwHbEVmsS0czzeAw5INwZxvsriQzYJllyIk").sheet1
    
    try:
        # Get all values as a list of lists
        all_values = sheet.get_all_values()
        if not all_values or len(all_values) < 2:
            print("No data found in sheet or insufficient rows")
            return
        
        # Headers are in row 2 (index 1)
        headers = all_values[1]
        processed_headers = []
        
        print(f"📋 Raw headers from row 2: {headers}")
        
        for i, header in enumerate(headers):
            if header.strip():
                processed_headers.append(header.strip())
            else:
                processed_headers.append(f"empty_column_{i+1}")
        
        print(f"📋 Processed headers: {processed_headers}")
        
        # Count total rows in sheet
        total_sheet_rows = len(all_values) - 2  # Exclude header rows
        print(f"📊 Total data rows in sheet: {total_sheet_rows}")
        
        # Process all rows including empty ones for debugging
        all_rows = []
        empty_rows = 0
        
        for row_idx, row_values in enumerate(all_values[2:], start=3):  # Start=3 for actual row numbers
            row = {}
            for header, value in zip(processed_headers, row_values):
                row[header] = value.strip() if value else ""
            
            # Check if row is completely empty
            if not any(cell.strip() for cell in row_values if cell):
                empty_rows += 1
                print(f"❌ Row {row_idx}: EMPTY ROW - Skipped")
                continue
                
            all_rows.append((row_idx, row))
        
        print(f"📊 Empty rows skipped: {empty_rows}")
        print(f"📊 Non-empty rows to process: {len(all_rows)}")
            
    except Exception as e:
        print(f"Error processing sheet data: {e}")
        return
    
    tz = pytz.timezone("Africa/Casablanca")
    
    # Detailed tracking
    created_count = 0
    skipped_no_timestamp = 0
    skipped_bad_timestamp = 0
    skipped_duplicates = 0
    skipped_other_errors = 0
    
    print("\n" + "="*50)
    print("DETAILED ROW-BY-ROW ANALYSIS")
    print("="*50)
    
    for row_idx, row in all_rows:
        print(f"\n🔄 Processing Sheet Row {row_idx}")
        
        # Show some key fields for context
        key_fields = {}
        for key in ['timestamp', 'Timestamp', 'nom_chauffeur', 'NOM DE CHAUFFEUR', 'client', 'CLIENTS']:
            if row.get(key):
                key_fields[key] = row.get(key)
        print(f"   Key fields: {key_fields}")
        
        # Check for timestamp - this is crucial
        timestamp_str = row.get("timestamp") or row.get("Timestamp")
        
        # Also check for other possible timestamp column names
        if not timestamp_str:
            # Check all columns that might contain timestamps
            possible_timestamp_cols = [col for col in processed_headers if 'time' in col.lower() or 'date' in col.lower()]
            print(f"   🔍 No 'timestamp'/'Timestamp' found. Checking other date columns: {possible_timestamp_cols}")
            
            for col in possible_timestamp_cols:
                if row.get(col) and ('/' in row.get(col) or ':' in row.get(col)):
                    timestamp_str = row.get(col)
                    print(f"   ✅ Found timestamp in column '{col}': {timestamp_str}")
                    break
        
        if not timestamp_str:
            print(f"   ❌ SKIPPED: No timestamp found")
            print(f"   📝 Available columns with data: {[k for k, v in row.items() if v]}")
            skipped_no_timestamp += 1
            continue

        print(f"   ⏰ Timestamp: '{timestamp_str}'")

        try:
            # Try multiple timestamp formats - American and European formats
            aware_dt = None
            formats_to_try = [
                "%m/%d/%Y %H:%M:%S",  # 6/30/2025 11:40:55 (American format)
                "%m/%d/%Y %H:%M",     # 6/30/2025 11:40 (American format without seconds)
                "%d/%m/%Y %H:%M:%S",  # 18/07/2025 15:30:51 (European format)
                "%d/%m/%Y %H:%M",     # 18/07/2025 15:30 (European format without seconds)
                "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d %H:%M",
                "%d-%m-%Y %H:%M:%S",
                "%d-%m-%Y %H:%M",
            ]
            
            for fmt in formats_to_try:
                try:
                    naive_dt = datetime.strptime(timestamp_str, fmt)
                    aware_dt = tz.localize(naive_dt)
                    print(f"   ✅ Timestamp parsed with format '{fmt}': {aware_dt}")
                    break
                except:
                    continue
            
            if not aware_dt:
                print(f"   ❌ SKIPPED: Could not parse timestamp '{timestamp_str}' with any format")
                skipped_bad_timestamp += 1
                continue
                
        except Exception as e:
            print(f"   ❌ SKIPPED: Timestamp error: {e}")
            skipped_bad_timestamp += 1
            continue

        # Check for duplicates
        if Livraison.objects.filter(timestamp=aware_dt).exists():
            print(f"   ⚠️ SKIPPED: Duplicate timestamp found in database")
            skipped_duplicates += 1
            continue

        # Try to create the record
        try:
            mapped_data = {
                'timestamp': aware_dt,
                'nom_chauffeur': row.get("nom_chauffeur") or row.get("NOM DE CHAUFFEUR") or "",
                'client': row.get("clients") or row.get("CLIENTS") or "",
                'chargements': row.get("LIEU DE CHARGE") or "",
                'date_chargement': parse_date_flexible(row.get("DATE CHARGE")),
                'dechargement': row.get("LIEU DE DECHARGE") or "",
                'bon_livraison': row.get("bon_livraison") or row.get("BON DE LIVRAISON") or "",
                'tarif': row.get("PRIX DE VOYAGE") or "",
                'deplacement': parse_decimal(row.get("deplacement") or row.get("DEPLACEMENT")),
                'avance': parse_decimal(row.get("avance")),
                'charge_variable': parse_decimal(row.get("CHARGE VARIABLE")),
                'prix_cv': parse_decimal(row.get("prix de ch.v") or row.get("PRIX DE CH.V")),
                'operateur': row.get("operateur") or row.get("OPERATEUR") or "",
                'commercial': row.get("commercial") or row.get("COMMERCIAL") or "",
                'ice': row.get("ice") or row.get("ICE") or "",
                'qte': parse_int(row.get("qte") or row.get("QTE")),
                'nom_destinataire': row.get("nom_destinataire") or row.get("NOM DE DESTINATAIRE") or "",
                'facturation': row.get("facturation") or row.get("Facturation") or ""
            }
            
            # Show non-empty mapped fields
            non_empty = {k: v for k, v in mapped_data.items() if v and v != ""}
            print(f"   📝 Non-empty fields: {list(non_empty.keys())}")

            Livraison.objects.create(**mapped_data)
            print(f"   ✅ SUCCESS: Created successfully!")
            created_count += 1
            
        except Exception as e:
            print(f"   ❌ SKIPPED: Database error: {e}")
            skipped_other_errors += 1

    print("\n" + "="*50)
    print("FINAL IMPORT SUMMARY")
    print("="*50)
    print(f"📊 Total data rows in sheet: {total_sheet_rows}")
    print(f"📊 Empty rows skipped: {empty_rows}")
    print(f"📊 Rows processed: {len(all_rows)}")
    print(f"✅ Successfully created: {created_count}")
    print(f"❌ Skipped - No timestamp: {skipped_no_timestamp}")
    print(f"❌ Skipped - Bad timestamp: {skipped_bad_timestamp}")
    print(f"⚠️ Skipped - Duplicates: {skipped_duplicates}")
    print(f"❌ Skipped - Other errors: {skipped_other_errors}")
    print(f"📈 Success rate: {(created_count/len(all_rows)*100):.1f}%" if all_rows else "0%")

# Helper functions
def parse_date(date_str):
    if not date_str:
        return None
    formats_to_try = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]
    for fmt in formats_to_try:
        try:
            return datetime.strptime(date_str, fmt).date()
        except:
            continue
    return None

def parse_decimal(value):
    if not value:
        return 0.0
    try:
        return float(str(value).replace(",", "."))
    except:
        return 0.0

def parse_int(value):
    if not value:
        return 0
    try:
        return int(float(value))
    except:
        return 0
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        'importer/credentials/sheetimporterproject-8a59547703c3.json', scope)
    client = gspread.authorize(creds)

    sheet = client.open_by_key("1jl2rRck8m4TeSaV91JjJ89kbIxXfNIWshQf_lAwbRYw").sheet1
    
    # Handle duplicate empty headers by using get_all_values() instead
    try:
        # Get all values as a list of lists
        all_values = sheet.get_all_values()
        if not all_values or len(all_values) < 2:
            print("No data found in sheet or insufficient rows")
            return
        
        # Headers are in row 2 (index 1), not row 1 (index 0)
        headers = all_values[1]
        processed_headers = []
        
        for i, header in enumerate(headers):
            if header.strip():
                processed_headers.append(header.strip())
            else:
                # Assign unique name to empty headers
                processed_headers.append(f"empty_column_{i+1}")
        
        # Convert to list of dictionaries (like get_all_records would do)
        rows = []
        for row_values in all_values[2:]:  # Start from row 3 (index 2) - skip first 2 rows
            if not any(cell.strip() for cell in row_values if cell):  # Skip completely empty rows
                continue
                
            row = {}
            for header, value in zip(processed_headers, row_values):
                row[header] = value.strip() if value else ""
            rows.append(row)
            
    except Exception as e:
        print(f"Error processing sheet data: {e}")
        return
    
    # Print all sheet titles (keeping your original functionality)
    for sheet_item in client.openall():
        print(sheet_item.title)
    
    tz = pytz.timezone("Africa/Casablanca")
    
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
            nom_chauffeur=row.get("nom_chauffeur") or row.get("NOM DE CHAUFFEUR") or "",
            client=row.get("clients") or row.get("CLIENTS") or "",
            chargements=row.get("LIEU DE CHARGE") or "",
            date_chargement=parse_date(row.get("date_chargement") or row.get("DATE DE CHARGE")),
            dechargement=row.get("LIEU DE DECHARGE") or "",  # Fixed this mapping
            bon_livraison=row.get("bon_livraison") or row.get("BON DE LIVRAISON") or "",
            tarif=row.get("PRIX DE VOYAGE") or "",
            deplacement=parse_decimal(row.get("deplacement") or row.get("DEPLACEMENT")),
            avance=parse_decimal(row.get("avance")),
            charge_variable=parse_decimal(row.get("CHARGE VARIABLE")),
            prix_cv=parse_decimal(row.get("prix de ch.v") or row.get("PRIX DE CH.V")),
            operateur=row.get("operateur") or row.get("OPERATEUR") or "",
            commercial=row.get("commercial") or row.get("COMMERCIAL") or "",
            ice=row.get("ice") or row.get("ICE") or "",
            qte=parse_int(row.get("qte") or row.get("QTE")),
            nom_destinataire=row.get("nom_destinataire") or row.get("NOM DE DESTINATAIRE") or "",
            facturation=row.get("facturation") or row.get("Facturation") or ""
        )

    print("✅ Livraisons imported successfully!")
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        'importer/credentials/sheetimporterproject-8a59547703c3.json', scope)
    client = gspread.authorize(creds)

    sheet = client.open_by_key("1jl2rRck8m4TeSaV91JjJ89kbIxXfNIWshQf_lAwbRYw").sheet1
    
    # Handle duplicate empty headers by using get_all_values() instead
    try:
        # Get all values as a list of lists
        all_values = sheet.get_all_values()
        if not all_values:
            print("No data found in sheet")
            return
        
        # Process headers and assign names to empty ones
        headers = all_values[0]
        processed_headers = []
        
        print(f"📋 Original headers: {headers}")  # DEBUG: Print original headers
        
        for i, header in enumerate(headers):
            if header.strip():
                processed_headers.append(header.strip())
            else:
                # Assign unique name to empty headers
                processed_headers.append(f"empty_column_{i+1}")
        
        print(f"📋 Processed headers: {processed_headers}")  # DEBUG: Print processed headers
        
        # Convert to list of dictionaries (like get_all_records would do)
        rows = []
        for row_values in all_values[1:]:  # Skip header row
            if not any(cell.strip() for cell in row_values if cell):  # Skip completely empty rows
                continue
                
            row = {}
            for header, value in zip(processed_headers, row_values):
                row[header] = value.strip() if value else ""
            rows.append(row)
            
        print(f"📊 Total rows to process: {len(rows)}")  # DEBUG: Print total rows
        
        # DEBUG: Print first row data to see the structure
        if rows:
            print(f"📄 First row data: {rows[0]}")
            
    except Exception as e:
        print(f"Error processing sheet data: {e}")
        return
    
    # Print all sheet titles (keeping your original functionality)
    for sheet_item in client.openall():
        print(sheet_item.title)
    
    tz = pytz.timezone("Africa/Casablanca")
    created_count = 0
    skipped_count = 0
    
    for i, row in enumerate(rows):
        print(f"\n🔄 Processing row {i+1}: {dict(list(row.items())[:3])}...")  # Show first 3 fields
        
        # Check for timestamp - this is crucial
        timestamp_str = row.get("timestamp") or row.get("Timestamp")
        print(f"⏰ Timestamp found: '{timestamp_str}'")  # DEBUG
        
        if not timestamp_str:
            print(f"❌ Row {i+1}: No timestamp found, skipping")
            skipped_count += 1
            continue

        try:
            # Convert string "14/07/2025 13:24:11" → aware datetime
            naive_dt = datetime.strptime(timestamp_str, "%d/%m/%Y %H:%M:%S")
            aware_dt = tz.localize(naive_dt)
            print(f"✅ Timestamp parsed successfully: {aware_dt}")
        except Exception as e:
            print(f"❌ Row {i+1}: Timestamp format issue: {e}")
            skipped_count += 1
            continue

        # Check for duplicates
        if Livraison.objects.filter(timestamp=aware_dt).exists():
            print(f"⚠️ Row {i+1}: Duplicate found, skipping")
            skipped_count += 1
            continue

        # DEBUG: Print all the mapped values before creating
        mapped_data = {
            'timestamp': aware_dt,
            'nom_chauffeur': row.get("nom_chauffeur") or row.get("NOM DE CHAUFFEUR") or "",
            'client': row.get("clients") or row.get("CLIENTS") or "",
            'chargements': row.get("LIEU DE CHARGE") or "",
            'date_chargement': parse_date(row.get("date_chargement") or row.get("DATE DE CHARGE")),
            'dechargement': row.get("lieu de charge") or row.get("LIEU DE DECHARGE") or "",
            'bon_livraison': row.get("bon_livraison") or row.get("BON DE LIVRAISON") or "",
            'tarif': row.get("PRIX DE VOYAGE") or "",
            'deplacement': parse_decimal(row.get("deplacement") or row.get("DEPLACEMENT")),
            'avance': parse_decimal(row.get("avance")),
            'charge_variable': parse_decimal(row.get("CHARGE VARIABLE")),
            'prix_cv': parse_decimal(row.get("prix de ch.v") or row.get("PRIX DE CH.V")),
            'operateur': row.get("operateur") or row.get("OPERATEUR") or "",
            'commercial': row.get("commercial") or row.get("COMMERCIAL") or "",
            'ice': row.get("ice") or row.get("ICE") or "",
            'qte': parse_int(row.get("qte") or row.get("QTE")),
            'nom_destinataire': row.get("nom_destinataire") or row.get("NOM DE DESTINATAIRE") or "",
            'facturation': row.get("facturation") or row.get("Facturation") or ""
        }
        
        print(f"🎯 Mapped data for row {i+1}:")
        for key, value in mapped_data.items():
            if value:  # Only print non-empty values
                print(f"   {key}: {value}")

        try:
            Livraison.objects.create(**mapped_data)
            print(f"✅ Row {i+1}: Successfully created!")
            created_count += 1
        except Exception as e:
            print(f"❌ Row {i+1}: Failed to create Livraison: {e}")
            skipped_count += 1

    print(f"\n📊 Import Summary:")
    print(f"   ✅ Created: {created_count}")
    print(f"   ⚠️ Skipped: {skipped_count}")
    print(f"   📋 Total processed: {len(rows)}")
    print("✅ Livraisons import completed!")

    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        'importer/credentials/sheetimporterproject-8a59547703c3.json', scope)
    client = gspread.authorize(creds)

    sheet = client.open_by_key("1jl2rRck8m4TeSaV91JjJ89kbIxXfNIWshQf_lAwbRYw").sheet1
    
    # Handle duplicate empty headers by using get_all_values() instead
    try:
        # Get all values as a list of lists
        all_values = sheet.get_all_values()
        if not all_values:
            print("No data found in sheet")
            return
        
        # Process headers and assign names to empty ones
        headers = all_values[0]
        processed_headers = []
        
        for i, header in enumerate(headers):
            if header.strip():
                processed_headers.append(header.strip())
            else:
                # Assign unique name to empty headers
                processed_headers.append(f"empty_column_{i+1}")
        
        # Convert to list of dictionaries (like get_all_records would do)
        rows = []
        for row_values in all_values[1:]:  # Skip header row
            if not any(cell.strip() for cell in row_values if cell):  # Skip completely empty rows
                continue
                
            row = {}
            for header, value in zip(processed_headers, row_values):
                row[header] = value.strip() if value else ""
            rows.append(row)
            
    except Exception as e:
        print(f"Error processing sheet data: {e}")
        return
    
    # Print all sheet titles (keeping your original functionality)
    for sheet_item in client.openall():
        print(sheet_item.title)
    
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
            nom_chauffeur=row.get("nom_chauffeur") or row.get("NOM DE CHAUFFEUR") or "",
            client=row.get("clients") or row.get("CLIENTS") or "",
            chargements=row.get("LIEU DE CHARGE") or "",
            date_chargement=parse_date(row.get("date_chargement") or row.get("DATE DE CHARGE")),
            dechargement=row.get("lieu de charge") or row.get("LIEU DE DECHARGE") or "",
            bon_livraison=row.get("bon_livraison") or row.get("BON DE LIVRAISON") or "",
            tarif=row.get("PRIX DE VOYAGE") or "",
            deplacement=parse_decimal(row.get("deplacement") or row.get("DEPLACEMENT")),
            avance=parse_decimal(row.get("avance")),
            charge_variable=parse_decimal(row.get("CHARGE VARIABLE")),
            prix_cv=parse_decimal(row.get("prix de ch.v") or row.get("PRIX DE CH.V")),
            operateur=row.get("operateur") or row.get("OPERATEUR") or "",
            commercial=row.get("commercial") or row.get("COMMERCIAL") or "",
            ice=row.get("ice") or row.get("ICE") or "",
            qte=parse_int(row.get("qte") or row.get("QTE")),
            nom_destinataire=row.get("nom_destinataire") or row.get("NOM DE DESTINATAIRE") or "",
            facturation=row.get("facturation") or row.get("Facturation") or ""
        )

    print("✅ Livraisons imported successfully!")