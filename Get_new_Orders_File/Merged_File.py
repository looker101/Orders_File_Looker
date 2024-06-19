import pandas as pd
from File_Matrixify import getShopifyFile

def mergedFile():
    df = getShopifyFile()
# Devo inserire nel nuovo file (df) tutto ciò che c'è nelle colonne: Ordered to, status, invoice e note del vecchio status order
    status = pd.read_excel("Order_Status_2024.xlsx")

    # per effettuare il merge, imposto come interi i numeri di ordine, ma prima devo riempire le celle vuote
    status["Name"] = status["Name"].fillna(0)
    status["Name"] = status["Name"].astype(int)

    # imposto come indice dello stato ordini il numero degli ordini
    #status.set_index("Name", inplace = True)

    # provo ad effettuare il merge dal file di sinitra ovvero df
    file_merged = df.merge(status, on = "Name", how = "left")
    file_merged.rename(columns={
        "Created At_x": "Created At",
        "Customer_x":"Customer",
        "SKU_x": "SKU", 
        "Quantity_x": "Quantity", 
        "Tags_x": "Tags"
    }, inplace=True)
    file_merged = file_merged[["Name", "Created At", "Customer", "SKU", "Quantity","Tags", "Ordered to", "Status", "Invoice", "Note"]]
    file_merged = file_merged.sort_values(by = "Name", ascending = True)
    #file_merged
    #file_merged = file_merged[["Name"]]
    file_merged.to_excel("Status_OrdersMerged.xlsx", index = False)
    #file_merged.to_excel("G:\\Il mio Drive\\Status_orders.xlsx")

if __name__ == "__main__":
    try:
        mergedFile()
        print("Merge dei file completato con successo")
    except Exception as e:
        print(f"Errore durante il merge dei file: {e}")