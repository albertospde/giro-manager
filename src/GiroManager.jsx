"""
MACRO IMPORT PRENOTATO → SUPABASE
Converte il file pianificazione visite Messaggerie nel formato
per l'upload nella tabella prenotato di Supabase Giro Manager.

Uso: python import_prenotato.py <file_input.xlsx> <file_output.xlsx>
"""

import pandas as pd
import sys
import os

# ─── MAPPING GRUPPO CLIENTE → CODICE CANALE SUPABASE ─────────────────────────
MAPPING_GRUPPI = {
    4:  'INDIPENDENTI_ALTRE_CATENE',
    6:  'LIB_RELIGIOSE',
    8:  'MONDADORI',
    11: 'LIBRACCIO',
    18: 'LIB_RELIGIOSE',
    19: 'LIB_RELIGIOSE',
    22: 'INDIPENDENTI_ALTRE_CATENE',
    23: 'FELTRINELLI',
    24: 'INDIPENDENTI_ALTRE_CATENE',
    25: 'GROSSISTI',
    28: 'FASTBOOK',
    30: 'GROSSISTI',
    32: 'GIUNTI',
    33: 'ALTRI_ONLINE',
    34: 'MONDADORI',
    36: 'LIB_COOP',
    56: 'LIBRACCIO',
    57: 'LIB_RELIGIOSE',
    58: 'IBS',
    59: 'FELTRINELLI',
    60: 'INDIPENDENTI_ALTRE_CATENE',
    63: 'CENTROLIBRI',
    77: 'FELTRINELLI',
    80: 'MONDADORI',
    82: 'AMAZON',
    83: 'UBIK',
    88: 'LIBRACCIO',
    90: 'LIBRACCIO',
    91: 'LIBRACCIO',
    92: 'LIBRACCIO',
    94: 'GROSSISTI',
    # 0 e NaN → INDIPENDENTI (gestito sotto)
}

def get_canale(gruppo):
    """Ritorna il codice canale Supabase dato il gruppo cliente."""
    if pd.isna(gruppo) or gruppo == 0:
        return 'INDIPENDENTI_ALTRE_CATENE'
    return MAPPING_GRUPPI.get(int(gruppo), None)


def converti(input_file, output_file):
    print(f"Lettura: {input_file}")
    df = pd.read_excel(input_file)

    # Cerca le colonne necessarie (flessibile sui nomi)
    col_map = {}
    for col in df.columns:
        col_lower = col.strip().lower()
        if 'gruppo' in col_lower and 'cliente' in col_lower:
            col_map['gruppo'] = col
        elif 'ean' in col_lower or 'codice' in col_lower and 'articolo' in col_lower:
            col_map['ean'] = col
        elif 'qtà' in col_lower or 'pren' in col_lower or 'quantit' in col_lower:
            col_map['qta'] = col
        elif 'codice' in col_lower and 'cliente' in col_lower:
            col_map['cod_cliente'] = col
        elif 'nome' in col_lower and 'cliente' in col_lower:
            col_map['nome_cliente'] = col

    print(f"Colonne trovate: {col_map}")
    
    # Applica mapping canale
    df['canale_codice'] = df[col_map['gruppo']].apply(get_canale)
    
    # Segnala gruppi non mappati
    non_mappati = df[df['canale_codice'].isna()][col_map['gruppo']].value_counts()
    if len(non_mappati) > 0:
        print(f"\n⚠️  GRUPPI NON MAPPATI (esclusi dall'output):")
        for g, cnt in non_mappati.items():
            tot = int(df[df[col_map['gruppo']] == g][col_map['qta']].sum())
            print(f"   Gruppo {g}: {cnt} righe, {tot} copie")
    
    # Aggregazione per EAN + canale + cliente
    df_agg = df[df['canale_codice'].notna()].groupby(
        [col_map['ean'], 'canale_codice', col_map['cod_cliente'], col_map['nome_cliente']],
        as_index=False
    )[col_map['qta']].sum()

    df_agg.columns = ['ean', 'canale_codice', 'codice_cliente', 'nome_cliente', 'quantita']
    df_agg = df_agg[df_agg['quantita'] > 0]

    # Riepilogo per canale
    print(f"\n📊 RIEPILOGO PER CANALE:")
    riepilogo = df_agg.groupby('canale_codice')['quantita'].sum().sort_values(ascending=False)
    for canale, tot in riepilogo.items():
        print(f"   {canale:<30} {int(tot):>8}")
    print(f"\n   {'TOTALE':<30} {int(df_agg['quantita'].sum()):>8}")

    # Salva output
    df_agg.to_excel(output_file, index=False)
    print(f"\n✅ Output salvato: {output_file}")
    print(f"   {len(df_agg)} righe pronte per l'import")


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Uso: python import_prenotato.py <input.xlsx> <output.xlsx>")
        sys.exit(1)
    converti(sys.argv[1], sys.argv[2])
