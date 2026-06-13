import os
import pandas as pd
from datetime import datetime

BRUTES_DIR = "Brutes"
CONCAT_DIR = "Concat"

os.makedirs(CONCAT_DIR, exist_ok=True)

def load_file(path):
    if path.endswith(".csv"):
        return pd.read_csv(path)
    elif path.endswith(".xlsx"):
        return pd.read_excel(path)
    return None

def load_all():
    frames = []
    for f in os.listdir(BRUTES_DIR):
        full = os.path.join(BRUTES_DIR, f)
        df = load_file(full)
        if df is not None:
            frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

def build_concat():
    df = load_all()

    # DEDUPLICATION (tu peux ajuster les colonnes clés ici)
    df = df.drop_duplicates()

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_file = os.path.join(CONCAT_DIR, f"SMS_Concat_{ts}.xlsx")

    df.to_excel(out_file, index=False)

    return out_file