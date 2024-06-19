import pandas as pd

def getShopifyFile():
    df = pd.read_excel("ShopifyExport.xlsx")
    # elimino tutte le righe vuote facendo riferimento alla colonna SKU
    df.dropna(axis = 0, subset=["Line: SKU"], inplace = True)
    # Rinomino le colonne
    df.rename(columns = {
        "Billing: Name":"Customer",
        "Line: Product Tags":"Tags",
        "Line: SKU":"SKU",
        "Line: Quantity":"Quantity"
    }, inplace = True)
    
    # setto gli indici
    df.set_index("Name", inplace = True)
    #elimino le colonne inutili
    df = df[["Created At", "Customer", "SKU", "Quantity", "Tags"]]
    # la colonna quantità deve essere un numero intero
    df["Quantity"] = df["Quantity"].astype("int")
    df.index = df.index.str.replace("#","")
    df.index = df.index.astype("int")

    #DISPONIBILI SUBITO
    # per far sì che la funzione faccia il suo dovere, devo trasformare i valori della colonna Tags in stringhe
    # alla colonna Tags applico la funzione dispoSubito(Se nella cella è presente "available now" la cella dovrà avere il solo valore "Disponibili subito")
    def dispoSubito(x):
        if "available now," in x:
            return "Disponibili Subito"
        return ""
        
    df["Tags"] = df["Tags"].astype(str)
    df["Tags"] = df["Tags"].apply(dispoSubito)

    #CREATED AT
    # nella colonna "Created at" rimuovo l'ora e UTC lasciando solo la data
    # alla colonna Created at assegno la funzione remove in modo che le celle contengano soltanto la data dell'acquisto
    def remove(x):
        return x[:10]

    df["Created At"] = df["Created At"].apply(remove)
    df.to_excel("New_status_Order.xlsx")
    return df
    
    

if __name__ == "__main__":
    try:
        getShopifyFile()
        print("Cancellate le colonne inutili")
    except Exception as e:
        print(f"Qualcosa è andato storto {e}")