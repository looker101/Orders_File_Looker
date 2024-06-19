from File_Matrixify import getShopifyFile
from Merged_File import mergedFile
from formatting import formattingStatusOrders

if __name__ == "__main__":
    try:
        getShopifyFile()
        print("File Shopify elaborato con successo")
    except Exception as e:
        print(f"Errore durante l'elaborazione del file Shopify: {e}")

    try:
        mergedFile()
        print("Merge dei file completato con successo")
    except Exception as e:
        print(f"Errore durante il merge dei file: {e}")

    try:
        file_path = "Status_OrdersMerged.xlsx"
        formattingStatusOrders(file_path)
        print("Formattazione del file completata con successo")
    except Exception as e:
        print(f"Errore durante la formattazione del file: {e}")
