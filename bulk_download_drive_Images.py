from pydrive2.auth import GoogleAuth 
from pydrive2.drive import GoogleDrive
import os
import re
import concurrent.futures
from tqdm import tqdm

# Configuración optimizada
DOWNLOAD_DIR = os.path.expanduser("~/Descargas/DrivePhotos")
LINKS_FILE = os.path.expanduser("~/Descargas/drivelinks.txt")
CREDENTIALS_FILE = "env/mycreds.txt" #esto lo mantengo oculto porque es un OAUTH
MAX_WORKERS = 5  # Número de descargas concurrentes

# Crear directorio de descargas
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# Autenticación mejorada
def setup_auth():
    gauth = GoogleAuth()
    gauth.DEFAULT_SETTINGS['client_config_file'] = "env/client_secrets.json" #esto lo mantengo oculto porque es un OAUTH
    gauth.Scopes = ["https://www.googleapis.com/auth/drive.readonly"]
    gauth.settings['get_refresh_token'] = True
    gauth.settings['save_credentials'] = True
    gauth.settings['save_credentials_backend'] = 'file'
    gauth.settings['save_credentials_file'] = CREDENTIALS_FILE
    
    if os.path.exists(CREDENTIALS_FILE):
        gauth.LoadCredentialsFile(CREDENTIALS_FILE)
    
    if gauth.credentials is None:
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        gauth.Refresh()
    else:
        gauth.Authorize()
    
    gauth.SaveCredentialsFile(CREDENTIALS_FILE)
    return GoogleDrive(gauth)

# Función para descargar un archivo individual
def download_file(drive, file_id, download_dir):
    try:
        file = drive.CreateFile({'id': file_id})
        file.FetchMetadata()
        filename = os.path.join(download_dir, file['title'])
        
        if os.path.exists(filename):
            return f"⚠️ Existe: {file['title']}"
            
        file.GetContentFile(filename)
        return f"✅ Descargado: {file['title']}"
    except Exception as e:
        return f"❌ Error {file_id}: {str(e)}"

# Procesamiento principal
def main():
    drive = setup_auth()
    
    # Extraer IDs de archivos
    try:
        with open(LINKS_FILE, 'r') as f:
            content = f.read()
        file_ids = list(set(re.findall(r'id=([a-zA-Z0-9_-]+)', content)))  # Elimina duplicados
        
        if not file_ids:
            raise ValueError("No se encontraron IDs en el archivo")
    except Exception as e:
        print(f"❌ Error leyendo archivo: {str(e)}")
        return

    # Descarga concurrente
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = []
        for file_id in file_ids:
            futures.append(executor.submit(download_file, drive, file_id, DOWNLOAD_DIR))
        
        # Barra de progreso
        for future in tqdm(concurrent.futures.as_completed(futures), total=len(futures)):
            print(future.result())

    print(f"\n🎉 Descargas completadas en: {DOWNLOAD_DIR}")

if __name__ == "__main__":
    main()
