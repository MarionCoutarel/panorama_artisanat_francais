
import os
import pandas as pd
import plotly.express as px
import xlrd

filepath ="https://www.data.gouv.fr/storage/f/2014-01-13T18-09-10/artisans_etalab.xls"

def data_prep(filepath):
    # first sheet data prep
    # _________________________________
    df = pd.read_excel(filepath, sheet_name=wb_sheets[0], header=None, skiprows=[
                   0, 1, 2], usecols=[1, 2, 3, 4, 5, 6, 7, 8, 9])
    df = df.rename(columns={1: "region", 2: "num_dpt", 3: "dpt", 4: "nb_entreprise",
                        5: "total", 6: "Alimentation", 7: "Fabrication",
                        8: "Batiment", 9: "Services"})
    df = df.fillna(method="ffill")
    df["Autre"] = df["nb_entreprise"] - df["total"]
    df = df.drop(["test", "total"], axis=1)
    # second sheet data prep
    # __________________________________
    df1 = pd.read_excel(
        filepath, sheet_name=wb_sheets[1], usecols=[1, 2, 3, 4, 5])
    df1 = df1.rename(columns={'Unnamed: 1': "region",
                          'Unnamed: 2': "num_dpt",
                          'Unnamed: 3': "dpt"})
    # merging 2 excel sheets
    # _________________________________
    DF = df.merge(df1, on=["region", "num_dpt", "dpt"])
    # adding population info
    # ________________________________
    pop = pd.read_excel('sources\pop_wikipedia.xlsx')
    DF.merge(pop, how="left", left_on="num_dpt", right_on="code")
    DF.drop(["code", "departement", "surface(Km²)"], axis=1)
    # adding Tableau readable region
    # __________________________________
    reg = pd.read_excel('sources\dpt_region.xlsx')
    DF.merge(reg, how='left', left_on="num_dpt", right_on="N°")
    DF = DF.drop(["region", "N°", "Département"], axis=1)

    return DF.to_excel("entreprises_artisanales.xlsx",
             sheet_name="entreprises artisanales")