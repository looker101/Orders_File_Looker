from Combined_Orders import getFile
from formatting import formattingStatusOrders

if __name__ == "__main__":
    try:
        print("Eseguendo getFile()...")
        getFile()
        print("getFile() eseguito con successo")
    except Exception as e:
        print(f"Errore durante l'esecuzione di getFile(): {e}")

    try:
        print("Eseguendo formattingStatusOrders()...")
        formattingStatusOrders()
        print("formattingStatusOrders() eseguito con successo")
    except Exception as e:
        print(f"Errore durante l'esecuzione di formattingStatusOrders(): {e}")