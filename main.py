import pandas as pd

# Charger le fichier Excel d'entrée
df = pd.read_excel("4.xlsx")

# Nettoyer les noms de colonnes
df.columns = df.columns.str.strip()

# Renommer les colonnes de base
df = df.rename(columns={
    'CdPerm': 'CliNoRéférence',
    'NomComplet': 'CliNomComplet',
    'LblAbrgPrg': 'CliListe_des_programmes_détudes',
    'LblConcent': 'CliConcentration',
    'CdRegrLieuEnseiChoix': 'CliUQOCampus',
    'DateDip': 'CliListe_des_programmes_détudes_Fin',
    'AdrCourrielInst': 'CliCourriel_Professionnel',
    'AdrCourriel': 'CliCourriel_Personnel',
    'NoTelCompletRes': 'CliTél_Domicile',
    'NoTelCompletCel': 'CliTél_Mobile',
    'Mun': 'CliVille',
    'Prv': 'CliProvince',
    'CdPost': 'CliCodePostal',
    'CdSect': 'CliCodeDépartement',
    
})

# Séparer NomComplet → CliNom + CliPrénom
df[['CliNom', 'CliPrénom']] = df['CliNomComplet'].str.extract(r'^([^,]+),\s*(.*)$')

# Fusionner NoCiv + Rue → CliAdresse
df['CliAdresse'] = (
    df['NoCiv'].fillna('').astype(str).str.strip() + ' ' +
    df['Rue'].fillna('').astype(str).str.strip()
).str.strip()

# Corriger le nom du campus
df['CliUQOCampus'] = df['CliUQOCampus'].replace({
    'GATIN': 'Gatineau',
    'STJER': 'Saint-Jérôme',
    'MANIW': 'Maniwaki',
    'MTLAU': 'Mont-Laurier',
    'STTHE': 'Sainte-Thérèse',
    'UQAC': 'UQAC',
    # Ripon
})

# Forcer le code secteur à 'DI'
df['CliCodeSecteur'] = 'DI'

# Convertir DateDip → "année:mois:jour"
df['CliListe_des_programmes_détudes_Fin'] = pd.to_datetime(
    df['CliListe_des_programmes_détudes_Fin'],
    errors='coerce'
).dt.strftime('%Y-%m-%d')

# Extraire l'année dans une nouvelle colonne
df['CliAnnée_de_diplomation'] = df['CliListe_des_programmes_détudes_Fin'].str[:4]

# Liste des provinces canadiennes (français + abréviations)
provinces_canada = {
    'AB', 'BC', 'MB', 'NB', 'NL', 'NS', 'NT', 'NU', 'ON', 'PE', 'QC', 'SK', 'YT',
    'Alberta', 'Colombie-Britannique', 'Manitoba', 'Nouveau-Brunswick',
    'Terre-Neuve-et-Labrador', 'Nouvelle-Écosse', 'Territoires du Nord-Ouest',
    'Nunavut', 'Ontario', 'Île-du-Prince-Édouard', 'Québec', 'Saskatchewan', 'Yukon'
}

# Correspondance des départements
departements_map = {
    'ADM': 'Sciences administratives',
    'DOYEN': 'Doyen',
    'EDUC': 'Sciences de l\'éducation',
    'INFO': 'Informatique et ingénierie',
    'PMM': 'Pmm',
    'R.IND': 'Relations Industrielles',
    'S.COM': 'Sciences comptables',
    'S.HUM': 'Sciences humaines'
}

# Appliquer la correspondance vers une nouvelle colonne
df['CliDépartementService'] = df['CliCodeDépartement'].map(departements_map).fillna('')

# Déterminer le pays selon la province
df['CliPays'] = df.apply(
    lambda row: row['CliProvince'] if str(row['CliProvince']).strip() not in provinces_canada else 'Canada',
    axis=1
)

# Définir l’ordre final des colonnes
colonnes_finales = [
    'CliNoRéférence',
    'CliNom',
    'CliPrénom',
    'CliListe_des_programmes_détudes',
    'CliConcentration',
    'CliUQOCampus',
    'CliCodeSecteur',
    'CliListe_des_programmes_détudes_Fin',
    'CliAnnée_de_diplomation',
    'CliCourriel_Professionnel',
    'CliCourriel_Personnel',
    'CliTél_Domicile',
    'CliTél_Mobile',
    'CliAdresse',
    'CliVille',
    'CliProvince',
    'CliCodePostal',
    'CliPays',
    'CliDépartementService'
]

# Réordonner les colonnes et remplacer les NaN par des chaînes vides
df = df.reindex(columns=colonnes_finales).fillna('')

# Sauvegarder le fichier final
df.to_csv("ProDon.csv", sep=';', index=False, encoding='utf-8-sig')
print("Conversion terminée : fichier ProDon.csv généré.")
