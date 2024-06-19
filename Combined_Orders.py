import pandas as pd

def load_excel(file_path, index_col=None):
    return pd.read_excel(file_path, index_col=index_col)

def clean_data(df):
    df = df.dropna(axis=0, subset=["Line: SKU"])
    df = df.rename(columns={
        "Billing: Name": "Customer",
        "Line: Product Tags": "Tags",
        "Line: SKU": "SKU",
        "Line: Quantity": "Quantity"
    })
    df.set_index("Name", inplace=True)
    df = df[["Created At", "Customer", "SKU", "Quantity", "Tags"]]
    df.loc[:, "Quantity"] = df["Quantity"].astype("int")
    df.index = df.index.str.replace("#", "")
    df.index = df.index.astype("int")
    return df

def get_new_orders(matrixify_orders, orders_status):
    new_orders_index = matrixify_orders.index.difference(orders_status.index)
    return matrixify_orders.loc[new_orders_index]

def dispo_subito(tags):
    if "available now," in tags:
        return "Disponibili Subito"
    return ""

def apply_tags(df):
    df["Tags"] = df["Tags"].astype(str)
    df["Tags"] = df["Tags"].apply(dispo_subito)
    return df

def save_to_excel(df, file_path):
    df.to_excel(file_path)

def getFile():
    matrixify_orders = load_excel("Matrixify_Orders_Export.xlsx")
    matrixify_orders = clean_data(matrixify_orders)
    
    orders_status = load_excel("Order_Status_2024.xlsx", index_col="Name")
    new_orders = get_new_orders(matrixify_orders, orders_status)
    new_orders = apply_tags(new_orders)
    
    combined_orders = pd.concat([orders_status, new_orders], ignore_index=False)
    save_to_excel(combined_orders, "Order_Status_2024OK.xlsx")

if __name__ == "__main__":
    getFile()