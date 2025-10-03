import requests
import urllib3

# Disabilita il warning SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = "https://raw.githubusercontent.com/Invenit-dev/Eaton-Xmoem-Online/main/QT_ITA-Offerta_Xmoem_v6.py"
local_filename = "QT_ITA-Offerta_Xmoem_v6.py"

try:
    # Scarica il file ignorando la verifica SSL
    response = requests.get(url, verify=False)
    response.raise_for_status()

    # Salva localmente
    with open(local_filename, "w", encoding="utf-8") as f:
        f.write(response.text)

    # Esegui lo script scaricato
    exec(open(local_filename, encoding="utf-8").read())

except Exception as e:
    print(f"Errore durante l'aggiornamento o l'esecuzione: {e}")
