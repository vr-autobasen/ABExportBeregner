import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
import requests
from datetime import datetime
import socket
import time

def load_config():
    try:
        with open('config.txt', 'r', encoding='utf-8') as file:
            config = {}
            for line in file:
                if '=' in line:
                    key, value = line.strip().split('=')
                    config[key.strip()] = value.strip()
            return config
    except FileNotFoundError:
        raise Exception("config.txt fil ikke fundet i samme mappe som scriptet")


def update_from_github():
    try:
        import requests
        raw_url = "https://raw.githubusercontent.com/vr-autobasen/ABExportBeregner/refs/heads/main/ExportCalc_inkl_van.py"
        response = requests.get(raw_url)

        if response.status_code == 200:
            with open(__file__, 'w', encoding='utf-8') as file:
                file.write(response.text)
            print("Script opdateret fra GitHub")
        else:
            print(f"Kunne ikke hente opdateringer. Status kode: {response.status_code}")
    except Exception as e:
        print(f"Fejl ved opdatering: {e}")


# Google Sheets setup
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
config = load_config()
SERVICE_ACCOUNT_FILE = config['SERVICE_ACCOUNT_FILE']
KM_SPREADSHEET_ID = config['KM_SPREADSHEET_ID']
TAX_SPREADSHEET_ID = config['TAX_SPREADSHEET_ID']

def load_config():
    try:
        with open('config.txt', 'r', encoding='utf-8') as file:
            config = {}
            for line in file:
                if '=' in line:
                    key, value = line.strip().split('=')
                    config[key.strip()] = value.strip()
            return config
    except FileNotFoundError:
        raise Exception("config.txt fil ikke fundet i samme mappe som scriptet")



def get_sheets_service():
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            creds = service_account.Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE, scopes=SCOPES)
            service = build('sheets', 'v4', credentials=creds)
            return service.spreadsheets()
        except socket.error:
            if attempt < max_attempts - 1:
                time.sleep(2 ** attempt)
                continue
            raise

def fetch_basic_vehicle_data(registration_number, api_token):
    url = f"https://api.nrpla.de/{registration_number}"
    headers = {"Authorization": f"Bearer {api_token}"}
    try:
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        data = response.json()["data"]
        return {
            'fuel_efficiency': data.get('fuel_efficiency'),
            'fuel_type': data.get('fuel_type'),
            'registration_date': data.get('first_registration_date'),
            'model': data.get('model'),
            'version': data.get('version'),
            'brand': data.get('brand'),
            'type': data.get('type'),
            'total_weight': data.get('total_weight')
        }
    except Exception as e:
        raise Exception(f"Fejl ved hentning af basis køretøjsdata: {e}")

def update_km_data(sheets, handelspris, norm_km, current_km):
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            updates = [
                {'range': 'Ark1!E7', 'values': [[handelspris]]},
                {'range': 'Ark1!E8', 'values': [[norm_km]]},
                {'range': 'Ark1!E9', 'values': [[current_km]]}
            ]

            for update in updates:
                sheets.values().update(
                    spreadsheetId=KM_SPREADSHEET_ID,
                    range=update['range'],
                    valueInputOption='RAW',
                    body={'values': update['values']}
                ).execute()
            break
        except socket.error:
            if attempt < max_attempts - 1:
                time.sleep(2 ** attempt)
                continue
            raise


def fetch_evaluation_data(registration_number, api_token):
    url = f"https://api.nrpla.de/evaluations/{registration_number}"
    headers = {"Authorization": f"Bearer {api_token}"}
    try:
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        data = response.json()["data"][0]
        return {
            'retail_price': data.get('retail_price'),
            'evaluation': data.get('evaluation', 0),
            'registration_tax': data.get('registration_tax', 0)
        }
    except Exception as e:
        raise Exception(f"Fejl ved hentning af evaluerings data: {e}")

def calculate_vehicle_age(registration_date):
    current_date = datetime.now()
    reg_date = datetime.strptime(registration_date, "%Y-%m-%d")
    return (current_date - reg_date).days // 365

def find_trade_price_based_on_age(sheets, vehicle_age):
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            result = sheets.values().get(
                spreadsheetId=KM_SPREADSHEET_ID,
                range='Ark1!E19:I19'
            ).execute()
            values = result.get('values', [[]])[0]

            if vehicle_age < 1:
                trade_price = values[0]
                age_group = "0-1 år"
            elif 1 <= vehicle_age < 2:
                trade_price = values[1]
                age_group = "1-2 år"
            elif 2 <= vehicle_age < 3:
                trade_price = values[2]
                age_group = "2-3 år"
            elif 3 <= vehicle_age < 10:
                trade_price = values[3]
                age_group = "3-9 år"
            else:
                trade_price = values[4]
                age_group = "Over 10 år"

            return float(trade_price) * 1000, age_group
        except socket.error:
            if attempt < max_attempts - 1:
                time.sleep(2 ** attempt)
                continue
            raise

def update_co2_in_sheets(sheets, fuel_type, fuel_efficiency, registration_date, vehicle_type):
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            if isinstance(fuel_efficiency, str):
                fuel_efficiency_formatted = fuel_efficiency.replace(".", ".")
            else:
                fuel_efficiency_formatted = str(fuel_efficiency).replace(".", ".")

            reg_date = datetime.strptime(registration_date, "%Y-%m-%d")
            wltp_cutoff_date = datetime.strptime("2017-09-01", "%Y-%m-%d")
            co2_norm = "WLTP" if reg_date >= wltp_cutoff_date else "NEDC"

            updates = [
                {'range': 'Værktøj til CO2!C26', 'values': [[co2_norm]]},
                {'range': 'Værktøj til CO2!C27', 'values': [[fuel_type]]},
                {'range': 'Værktøj til CO2!C25', 'values': [[fuel_efficiency_formatted]]}
            ]

            for update in updates:
                sheets.values().update(
                    spreadsheetId=TAX_SPREADSHEET_ID,
                    range=update['range'],
                    valueInputOption='USER_ENTERED',
                    body={'values': update['values']}
                ).execute()

            result = sheets.values().get(
                spreadsheetId=TAX_SPREADSHEET_ID,
                range='Værktøj til CO2!C30'
            ).execute()
            co2_value = result.get('values', [[0]])[0][0]

            target_range = 'Brugte Varebiler!L23' if vehicle_type == "Varebil" else 'co2km01'
            sheets.values().update(
                spreadsheetId=TAX_SPREADSHEET_ID,
                range=target_range,
                valueInputOption='USER_ENTERED',
                body={'values': [[co2_value]]}
            ).execute()
            break
        except socket.error:
            if attempt < max_attempts - 1:
                time.sleep(2 ** attempt)
                continue
            raise

def update_vehicle_data(sheets, vehicle_type, total_weight, handelspris, new_price):
    if vehicle_type == "Varebil":
        weight_category = "over 3.000 kg og som enten er åben eller uden sideruder bag føresædet" if total_weight > 3000 else "Alle andre"
        updates = [
            {'range': 'Brugte Varebiler!L21', 'values': [[str(int(handelspris))]]},
            {'range': 'Brugte Varebiler!L22', 'values': [[str(int(new_price))]]},
            {'range': 'Brugte Varebiler!L27', 'values': [[weight_category]]}
        ]
    else:
        updates = [
            {'range': 'handelspris01', 'values': [[str(int(handelspris))]]},
            {'range': 'nypris01', 'values': [[str(int(new_price))]]}
        ]

    for update in updates:
        sheets.values().update(
            spreadsheetId=TAX_SPREADSHEET_ID,
            range=update['range'],
            valueInputOption='RAW',
            body={'values': update['values']}
        ).execute()

def get_export_tax(sheets, vehicle_type):
    tax_range = 'Brugte Varebiler!G32' if vehicle_type == "Varebil" else 'finalTax01'
    result = sheets.values().get(
        spreadsheetId=TAX_SPREADSHEET_ID,
        range=tax_range
    ).execute()
    return float(result.get('values', [[0]])[0][0])

def calculate_new_price(eval_data, manual_price=None):
    if manual_price is not None:
        try:
            return float(manual_price)
        except ValueError:
            raise Exception("Ugyldig manuel pris indtastet")

    if eval_data.get('retail_price'):
        return eval_data['retail_price']
    elif eval_data.get('evaluation') and eval_data.get('registration_tax'):
        return eval_data['evaluation'] + eval_data['registration_tax']
    else:
        return None


def log_to_file(registration_number, type, vehicle_info, new_price, export_tax, reduced_tax, handelspris_input, norm_km_input, current_km_input, sheet_handelspris, age_group):
    # Opret logs mappe hvis den ikke eksisterer
    if not os.path.exists('logs'):
        os.makedirs('logs')

    # Generer filnavn med dagens dato
    filename = f"logs/vehicle_export_log_{datetime.now().strftime('%Y-%m-%d')}.txt"

    # Få antal eksisterende entries i dagens log fil
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            entry_count = sum(1 for line in f if line.startswith('=== Log Entry'))
    except FileNotFoundError:
        entry_count = 0

    # Formater log entry
    log_entry = (
        f"\n=== Log Entry #{entry_count + 1} - {datetime.now().strftime('%H:%M:%S')} ===\n"
        f"1. Nummerplade: {registration_number}\n"
        f"2. Type: {type}\n"
        f"3. Køretøj: {vehicle_info}\n"
        f"4. Indtastet handelspris: {handelspris_input:,.2f} kr.\n"
        f"5. Norm kilometer: {norm_km_input:,} km\n"
        f"6. Aktuelle kilometer: {current_km_input:} km\n"
        f"7. Handelspris fra sheet: {sheet_handelspris:,.2f} kr. ({age_group})\n"
        f"8. Nypris: {new_price:,.2f} kr.\n"
        f"9. Eksportafgift: {export_tax:.2f} kr.\n"
        f"10. Eksportafgift efter reduktion: {reduced_tax:.2f} kr.\n"
        f"{'=' * 50}\n"
    )

    # Skriv til logfil
    with open(filename, 'a', encoding='utf-8') as f:
        f.write(log_entry)


def main():
    config = load_config()

    api_token = config['API_TOKEN']

    while True:
        try:
            sheets = get_sheets_service()

            # Spørg efter nummerplade
            registration_number = input("\nIndtast nummerplade (eller 'q' for at afslutte): ").strip()

            # Check om brugeren vil afslutte
            if registration_number.lower() == 'q':
                print("Afslutter programmet...")
                break
            print("Henter basis køretøjsdata...")
            basic_data = fetch_basic_vehicle_data(registration_number, api_token)
            vehicle_type = basic_data['type']
            total_weight = basic_data.get('total_weight', 0)

            print("Henter evalueringsdata...")
            eval_data = fetch_evaluation_data(registration_number, api_token)

            vehicle_age = calculate_vehicle_age(basic_data['registration_date'])
            print(f"Bilens alder: {vehicle_age} år")

            handelspris_input = float(input("Indtast handelsprisen: "))
            norm_km_input = float(input("Indtast norm km: "))
            current_km_input = float(input("Indtast bilens kørte kilometer: "))

            update_km_data(sheets, handelspris_input, norm_km_input, current_km_input)
            handelspris, age_group = find_trade_price_based_on_age(sheets, vehicle_age)
            print(f"Handelspris fra sheet: {handelspris} kr. for aldersgruppen {age_group}.")

            new_price = calculate_new_price(eval_data)
            if new_price is None:
                manual_price = input("Kunne ikke beregne nypris automatisk. Indtast manuel nypris: ")
                new_price = calculate_new_price(eval_data, manual_price)

            update_co2_in_sheets(sheets, basic_data['fuel_type'], basic_data['fuel_efficiency'],
                               basic_data['registration_date'], vehicle_type)

            update_vehicle_data(sheets, vehicle_type, total_weight, handelspris, new_price)

            export_tax = get_export_tax(sheets, vehicle_type)

            brand = basic_data.get('brand', 'N/A')
            model = basic_data.get('model', 'N/A')
            version = basic_data.get('version', 'N/A')
            fuel_type = basic_data.get('fuel_type', 'N/A')
            vehicle_info = f"{brand} {model} {version} {fuel_type}"

            print(f"\nType: {vehicle_type}")
            if vehicle_type == "Varebil":
                print(f"Totalvægt: {total_weight} kg")
            print(f"Køretøj: {vehicle_info}")
            print(f"Nypris: {new_price:,.2f} kr.")
            print(f"Eksportafgift: {export_tax:.2f} kr.")
            reduced_tax = (export_tax * 0.85 - 3000) if export_tax > 50000 else export_tax - 11000
            print(f"Eksportafgift efter reduktion: {reduced_tax:.2f} kr.")
            log_to_file(registration_number, vehicle_type, vehicle_info, new_price, export_tax, reduced_tax,
                        handelspris_input, norm_km_input, current_km_input, handelspris, age_group)



        except Exception as e:
            print(f"Fejl: {e}")
            time.sleep(2)
            continue

if __name__ == "__main__":
    update_from_github()
    main()