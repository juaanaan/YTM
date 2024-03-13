import requests
import re
import time
import json
import sys
from cryptography.fernet import Fernet
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta
from urllib.parse import quote_plus

def writeCSC(service_sheets, spreadsheet_id, index):
    service_sheets.spreadsheets().values().update(spreadsheetId=spreadsheet_id,
                                range='Hoja 1!G' + str((index + 2)),
                                valueInputOption='USER_ENTERED',
                                body={'values': [['CSC']]}
                            ).execute()  

def writeCCC(service_sheets, spreadsheet_id, index):
    service_sheets.spreadsheets().values().update(
                            spreadsheetId=spreadsheet_id,
                            range='Hoja 1!G' + str((index + 2)),
                            valueInputOption='USER_ENTERED',
                            body={'values': [['CCC']]}
                        ).execute()  

def writeHash(service_sheets, spreadsheet_id, index, existingHash):
    service_sheets.spreadsheets().values().update(
                            spreadsheetId=spreadsheet_id,
                            range='Hoja 1!N' + str((index + 2)),
                            valueInputOption='USER_ENTERED',
                            body={'values': [[existingHash]]}
                        ).execute()
    
def cargar_clave():
    with open('/Users/Juaanaan_/Desktop/Yourtaximallorca/key.key', 'rb') as f:
        return f.read()

# Función para desencriptar el contenido del archivo JSON
def desencriptar_json(clave, archivo_encriptado):
    cipher_suite = Fernet(clave)
    json_encriptado = open(archivo_encriptado, 'rb').read()
    json_desencriptado = cipher_suite.decrypt(json_encriptado)
    return json_desencriptado.decode()


def is_code_in_sheets(code, data):
    """Función para verificar si un código de reserva ya existe en Google Sheets."""
    for row in data:
        if row[0] == code:
            return True
    return False

def iso8601_to_dd_mm_yyyy_hh_mm_ss(iso8601_string):
    """Función para convertir una cadena ISO8601 a formato de fecha y hora legible."""
    try:
        dt_object = datetime.fromisoformat(iso8601_string)
        formatted_datetime = dt_object.strftime("%d-%m-%Y %H:%M:%S")
        return formatted_datetime
    except ValueError:
        return "Error: Formato de fecha y hora no válido"

def generate_event_summary(code, flight_number, pickup_address, dropoff_address, pickup_time, client_name, email, phone, passengerCount, luggageCount, driverCode, pickup_maps_link, addOns, comments):
    """Función para generar el resumen del evento."""
    pickup_time = iso8601_to_dd_mm_yyyy_hh_mm_ss(pickup_time)
    summary = f"*Código de Reserva:* {code}\n\n"
    summary += f"*Link conductor:* https://staging.transferz.taxi/{driverCode}\n\n"
    summary += f"*Número de Vuelo:* {flight_number}\n\n"
    summary += f"*Dirección de Recogida:* {pickup_address}\n\n"
    summary += f"*Dirección de Destino:* {dropoff_address}\n\n"
    summary += f"*Fecha de Recogida:* {pickup_time}\n\n"
    summary += f"*Cliente:* {client_name}\n\n"
    summary += f"*Email Cliente:* {email}\n\n"
    summary += f"*Teléfono Cliente:* {phone}\n\n"
    summary += f"*Número de Pasajeros:* {passengerCount}\n\n"
    summary += f"*Número de Maletas:* {int(luggageCount)}\n\n"
    
    if comments:
        summary += f"*Comentarios:* {comments}\n\n"

    if addOns:
        summary += "*Extras:*\n"
        for addon in addOns:
            summary += f"- {addon}\n"
    summary += f"\n*Ruta Google Maps:* {pickup_maps_link}\n"
    return summary

def generate_whatsapp_link(message):
    """Función para generar un enlace de WhatsApp con un mensaje predefinido."""
    # Formato del enlace de WhatsApp
    return f"https://wa.me/?text={quote_plus(message)}"

def airport(pickup_address, dropoff_address, vehicleCategory):
    """Función para identificar y cambiar el nombre de direcciones de aeropuerto."""
    if pickup_address == "Palma de Mallorca Airport (PMI), 07611 Palma, Illes Balears, Spain":
        pickup_address = "Aeropuerto"
        color_id = 2
    elif dropoff_address == "Palma de Mallorca Airport (PMI), 07611 Palma, Illes Balears, Spain":
        dropoff_address = "Aeropuerto"
        color_id = 11
    
    if vehicleCategory == "MINIVAN" or vehicleCategory == "MINIBUS":
        color_id = 5

    return pickup_address, dropoff_address, color_id

def generate_google_maps_link(origin, destination):
    """Función para generar el enlace de Google Maps con las direcciones de origen y destino."""
    origin_encoded = quote_plus(origin)
    destination_encoded = quote_plus(destination)
    return f"https://www.google.com/maps/dir/?api=1&origin=current+location&destination={destination_encoded}&waypoints={origin_encoded}&travelmode=driving"


def create_event(service, calendar_id, code, flight_number, pickup_address, dropoff_address, pickup_time, client_name, email, phone, passengerCount, luggageCount, pickup_maps_link, price, whatsapp_link, color_id, addOns, comments):
    """Función para crear un evento en Google Calendar."""
    # Creación de la hora de finalización para evento en Google Calendar (+1 hora)
    pickup_time_dt = datetime.fromisoformat(pickup_time)
    pickup_time_dt_f = pickup_time_dt + timedelta(hours=1)
    pickup_time_f = pickup_time_dt_f.isoformat()

    # Construcción del contenido del evento
    event_description = f"""\
<b>Código de Reserva:</b> {code}<br>
<b>Precio (con IVA):</b> {price} €<br>
<b>Número de Vuelo:</b> {flight_number}<br>
<b>Dirección de Recogida:</b> {pickup_address}<br>
<b>Dirección de Destino:</b> {dropoff_address}<br>
<b>Ruta:</b> <a href="{pickup_maps_link}" target="_blank">Ver ruta en Google Maps</a><br>
<b>Whatsapp:</b> <a href="{whatsapp_link}" target="_blank">Enviar información</a><br>
<b>Fecha de Recogida:</b> {iso8601_to_dd_mm_yyyy_hh_mm_ss(pickup_time)}<br>
<b>Cliente:</b> {client_name}<br>
<b>Email Cliente:</b> {email}<br>
<b>Teléfono Cliente:</b> {phone}<br>
<b>Número de Pasajeros:</b> {passengerCount}<br>
<b>Número de Maletas:</b> {int(luggageCount)}<br>
"""

    if addOns:
        event_description += "<b>Extras:</b>\n"
        event_description += "\n".join([f"- {addon}" for addon in addOns])
    
    # Añadir sección de comentarios solo si hay comentarios
    if comments:
        event_description += f"<br><b>Comentarios:</b> {comments}"

    # Agregar espaciado entre los elementos del evento
    event_description = event_description.replace("<br>", "<br><br>")

    

    # Creación del objeto de evento
    event = {
        'summary': f'Transferz: {code}',
        'description': event_description,
        'start': {
            'dateTime': pickup_time,
            'timeZone': 'Europe/Madrid',
        },
        'end': {
            'dateTime': pickup_time_f,
            'timeZone': 'Europe/Madrid',
        },
        'colorId': color_id  # Especificar el color deseado
    }

    # Inserción del evento en Google Calendar
    event = service.events().insert(calendarId=calendar_id, body=event).execute()
    print('Evento creado: %s' % (event.get('htmlLink')))
    sys.stdout.flush()



def write_to_google_sheets(service, spreadsheet_id, data):
    """Función para escribir datos en Google Sheets."""
    range_name = 'Hoja 1!A2'  # Especifica el rango donde deseas insertar los datos
    value_input_option = 'USER_ENTERED'

    values = []
    for row in data:
        # Extraemos los datos de la fila
        code, pickup_time, entry_or_exit, pickup_time_hour, pickup_address, dropoff_address, price, hash, status = row

        if pickup_address == "Aeropuerto":
            non_airport_address = dropoff_address
        else:
            non_airport_address = pickup_address

        pattern = r"\b07\d{3}\b\s*(.*?)[,-]"

        # Buscar el patrón en la dirección de dropoff
        match = re.search(pattern, non_airport_address)

        if match:
            # Extraer la parte después del código postal
            extracted_address = match.group(1).strip()
            print("Parte después del código postal:", extracted_address)
            sys.stdout.flush()
            
            # Comprobamos si la reserva ha sido cancelada
            if status == "CANCELLED_FREE":
                status = "CSC"
                values.append([code, pickup_time, entry_or_exit, pickup_time_hour, extracted_address, price, status, "", "", "", "", "", "", hash])
            elif status == "CANCELLED_WITH_COSTS":
                    status = "CCC"
                    values.append([code, pickup_time, entry_or_exit, pickup_time_hour, extracted_address, price, status, "", "", "", "", "", "", hash])
            else:
                values.append([code, pickup_time, entry_or_exit, pickup_time_hour, extracted_address, price, "", "", "", "", "", "", "", hash])

        else:
            print("No se encontró un código postal que comience con '07' o la dirección no sigue el formato esperado.")
            sys.stdout.flush()
            
            # Comprobamos si la reserva ha sido cancelada
            if status == "CANCELLED_FREE":
                status = "CSC"
                values.append([code, pickup_time, entry_or_exit, pickup_time_hour, non_airport_address, price, status, "", "", "", "", "", "", hash])
            elif status == "CANCELLED_WITH_COSTS":
                    status = "CCC"
                    values.append([code, pickup_time, entry_or_exit, pickup_time_hour, non_airport_address, price, status, "", "", "", "", "", "", hash])
            else:
                values.append([code, pickup_time, entry_or_exit, pickup_time_hour, non_airport_address, price, "", "", "", "", "", "", "", hash])
            
    body = {
        'values': values
    }
    
    try:
        # Insertar datos en Google Sheets
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption=value_input_option,
            body=body
        ).execute()
        print('Datos insertados correctamente en Google Sheets.')
        sys.stdout.flush()

        # Obtener la cantidad de filas y columnas con datos
        sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        num_rows = sheet_metadata['sheets'][0]['properties']['gridProperties']['rowCount']
        num_cols = sheet_metadata['sheets'][0]['properties']['gridProperties']['columnCount']

        # Construir el rango dinámico
        sort_range = {
            'sheetId': 0,
            'startRowIndex': 1,
            'startColumnIndex': 0,
            'endRowIndex': num_rows,
            'endColumnIndex': num_cols
        }

        # Llamar al script de Google Sheets para ordenar todas las filas por fecha
        request = {
            'sortRange': {
                'range': sort_range,
                'sortSpecs': [
                    {
                        'dimensionIndex': 1,  # Ordenar por la segunda columna (B)
                        'sortOrder': 'DESCENDING'  # Puedes cambiar a 'DESCENDING' si lo prefieres
                    }
                ]
            }
        }
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': [request]}
        ).execute()
    except Exception as e:
        print('Error al insertar datos en Google Sheets:', e)
        sys.stdout.flush()



def main():
    """
    Función principal del script.
    """
    # Credenciales del servicio de Google Calendar
    credentials = service_account.Credentials.from_service_account_file(
        '//Users/Juaanaan_/Desktop/Yourtaximallorca/ytm-api-414822-2da39b0051e3.json',
        scopes=['https://www.googleapis.com/auth/calendar']
    )

    # Construcción del servicio de Google Calendar
    service = build('calendar', 'v3', credentials=credentials)

    # Credenciales del servicio de Google Sheets
    credentials_sheets = service_account.Credentials.from_service_account_file(
        '/Users/Juaanaan_/Desktop/Yourtaximallorca/ytm-api-414822-2da39b0051e3.json',
        scopes=['https://www.googleapis.com/auth/spreadsheets']
    )

    # Construcción del servicio de Google Sheets
    service_sheets = build('sheets', 'v4', credentials=credentials_sheets)

    # ID de la hoja de cálculo
    spreadsheet_id = '1qqFQ06lb3LFXc3wF4gMZ5BapgPI9V523xUIMGsadx6w'  # Reemplazar con el ID de tu hoja de cálculo

    # Obtener la hora actual menos una hora y diez minutos
    current_time_minus_one_hour = datetime.now() - timedelta(hours=1, minutes=20)
    formatted_time = current_time_minus_one_hour.strftime('%Y-%m-%dT%H:%M:%S')
    print(formatted_time)
    sys.stdout.flush()
    
    clave = cargar_clave()

    # Desencriptar el archivo JSON
    json_desencriptado = desencriptar_json(clave, '/Users/Juaanaan_/Desktop/Yourtaximallorca/config_encrypted.json')

    # Convertir el JSON desencriptado a un diccionario
    config = json.loads(json_desencriptado)

    api_key = config["API_KEY"]
    # Construir la URL de la consulta con la hora actual menos una hora y diez minutos
    url = f"https://warpdrive.staging.transferz.com/transfercompanies/journeys?excludedStatuses=COMPLETED&excludedStatuses=PENDING&updatedAfter={formatted_time}&page=0&size=20"

    headers = {
        "accept": "application/json",
        "X-API-Key": api_key
    }

    # Obtener datos de la primera columna de la hoja de cálculo para verificar códigos de reserva existentes
    result = service_sheets.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range='Hoja 1!A2:A').execute()
    values = result.get('values', [])
    
    # Extraer los códigos de reserva existentes de la hoja de cálculo
    existing_codes = [row[0] for row in values if row]

    response = requests.get(url, headers=headers, timeout=10)  # Tiempo de espera de 10 segundos

    if response.status_code == 200:
        data = response.json()
        results = data.get("results", [])
        if results:
            for result in results: # Obtención de datos para transfer
                code = result.get("code")
                hash = result.get("hash")
                status = result.get("status")

                # Verificar si el código de reserva ya existe en la hoja de cálculo
                if code in existing_codes:
                    # Comprobamos si el hash ha cambiado
                    index = existing_codes.index(code)
                    existingHash = service_sheets.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range='Hoja 1!N' + str((index + 2))).execute()
                    
                    if existingHash == hash:
                        print(f"El servicio con el código de reserva {code} ya está en la hoja de cálculo. Descartando...")
                        sys.stdout.flush()
                        continue
                    
                    # En caso de que no coincidan los hash, querrá decir que ha habido una actualización en la reserva
                    updated = True

                    # Comprobamos si la reserva ha sido cancelada
                    
                    if status == "CANCELLED_FREE":
                        writeCSC(service_sheets, spreadsheet_id, index)    
                        writeHash(service_sheets, spreadsheet_id, index, existingHash)             
                        continue
                    elif status == "CANCELLED_WITH_COSTS":          
                        writeCCC(service_sheets, spreadsheet_id, index)
                        writeHash(service_sheets, spreadsheet_id, index, existingHash)
                        continue
                
                flight_number = result.get("travellerInfo", {}).get("flightNumber")
                pickup_address = result.get("pickup", {}).get("resolvedAddress")
                
                if pickup_address == None:
                    pickup_address = result.get("pickup", {}).get("bookerEnteredAddress")
                
                dropoff_address = result.get("dropoff", {}).get("resolvedAddress")
                pickup_time = iso8601_to_dd_mm_yyyy_hh_mm_ss(result.get("pickupTime", {}).get("localTime"))
                client_name = result.get("travellerInfo", {}).get("firstName") + " " + result.get("travellerInfo", {}).get("lastName")
                email = result.get("travellerInfo", {}).get("email")
                phone = result.get("travellerInfo", {}).get("phone")
                passengerCount = result.get("travellerInfo", {}).get("passengerCount")
                luggageCount = result.get("travellerInfo", {}).get("luggageCount")
                comments = result.get("travellerInfo", {}).get("driverComments")
                price  = result.get("fareSummary",{}).get("includingVat")
                driverCode = result.get("driverCode",{})
                vehicleCategory = result.get("vehicleCategory",{})
                addOns = result.get("addOns", [])               
                # Creación ruta Google Maps

                pickup_maps_link = generate_google_maps_link(pickup_address, dropoff_address)

                # Cambio de nombre a Aeropuerto y asignación del color del evento

                pickup_address, dropoff_address, color_id = airport(pickup_address, dropoff_address, vehicleCategory)

                if status != "CANCELLED_FREE" and status != "CANCELLED_WITH_COSTS":
                    # Creación de Whatsapp

                    event_summary = generate_event_summary(code, flight_number, pickup_address, dropoff_address, result.get("pickupTime", {}).get("localTime"), client_name, email, phone, passengerCount, luggageCount, driverCode, pickup_maps_link, addOns, comments)
                    whatsapp_link = generate_whatsapp_link(event_summary)

                    # Creación del nuevo evento

                    create_event(service, 'juanangonzcomas@gmail.com', code, flight_number, pickup_address, dropoff_address, result.get("pickupTime", {}).get("localTime"), client_name, email, phone, passengerCount, luggageCount, pickup_maps_link, price, whatsapp_link, color_id, addOns, comments)

                rows_to_append = []
                
                # Define el valor de la columna "si es una entrada o una salida" según el origen
                entry_or_exit = "ENTRADA" if pickup_address == "Aeropuerto" else "SALIDA"

                # Extraemos la hora de la hora de recogida
                pickup_day, pickup_hour = pickup_time.split()

                # Agregamos los datos a rows_to_append
                
                rows_to_append.append([code, pickup_day, entry_or_exit, pickup_hour, pickup_address, dropoff_address, price, hash, status])

                write_to_google_sheets(service_sheets, spreadsheet_id, rows_to_append)

    else:
        print("La solicitud no fue exitosa. Código de estado:", response.status_code)
        sys.stdout.flush()

if __name__ == "__main__":
    while True:
        try:
            main()
            # Esperar 5 minutos antes de la próxima ejecución
            time.sleep(300)  # 300 segundos = 5 minutos
        except Exception as e :
            print("Error al ejecutar Script:")
            print("")
            print(e)
            continue
