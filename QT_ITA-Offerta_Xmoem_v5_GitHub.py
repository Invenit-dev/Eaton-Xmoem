import urllib.request
import ssl

url = "https://raw.githubusercontent.com/Invenit-dev/Eaton-Xmoem-Online/main/QT_ITA-Offerta_Xmoem_v6.py"
local_filename = "QT_ITA-Offerta_Xmoem_v6.py"

# Crea un contesto SSL che ignora la verifica del certificato
ssl_context = ssl._create_unverified_context()

try:
    with urllib.request.urlopen(url, context=ssl_context) as response:
        content = response.read().decode("utf-8")

    with open(local_filename, "w", encoding="utf-8") as f:
        f.write(content)

    exec(open(local_filename, encoding="utf-8").read())

except Exception as e:
    print(f"Errore durante l'aggiornamento o l'esecuzione: {e}")
