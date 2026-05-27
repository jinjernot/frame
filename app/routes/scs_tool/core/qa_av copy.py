#Importar librería
import pandas as pd
import sys

def av_check(df):

    #Cargar la informacion de los archivos en data frames
    SkuAcc=pd.read_excel(df, sheet_name="SKU Accuracy")
    MS4_report=pd.read_excel(df, sheet_name="ms4")

    #Eliminar expacios excesivos y separar los SKUs y AVs de MS4 de sus regiones
    ms4sku_df=pd.DataFrame(MS4_report["SKU                                     "].str.split('#').str[0])
    ms4av_df=pd.DataFrame(MS4_report["SKU AV                                  "].str.split('#').str[0])
    ms4sku_df=ms4sku_df.replace(r"^ +| +$", r"", regex=True)
    ms4av_df=ms4av_df.replace(r"^ +| +$", r"", regex=True)

    #Validacion de datos de ambos dataframes ya limpios
    MS4_len=len(ms4sku_df.drop_duplicates())
    SKUAcc_len=len(SkuAcc["SKU"].drop_duplicates())
    if (MS4_len!=SKUAcc_len):
        print("Data of one of both tabs are corrupted, please verify the file and try again")
        sys.exit()

    #Crear nuevos data frames para el análisis
    #DataFrame de SCS (Sku Accuracy)
    SKuAcc_df=pd.DataFrame(SkuAcc["Component"]).join(pd.DataFrame(SkuAcc["ComponentGroup"])).join(SkuAcc["SKU"])
    SKuAcc_df=SKuAcc_df.rename(columns={"SKU":"SCS_SKU","Component":"Component_SCS"})
    #DataFrame de BOM (MS4)
    MS4_df=ms4sku_df.join(ms4av_df)
    MS4_df=MS4_df.rename(columns={"SKU                                     ":"BOM_SKU","SKU AV                                  ":"Component_BOM"})

    #Crear nuevo data frame, buscando si un AV está asociado a X SKU
    lookup_df=SKuAcc_df.merge(
        MS4_df,
        left_on="Component_SCS",
        right_on="Component_BOM",
        how="left")

    #Filtrar AVs que no existan en algun SKU
    filter_df=lookup_df[lookup_df["BOM_SKU"].isnull()]

    #DataFrame que guarda los resultados del filtro
    final_df=pd.DataFrame(filter_df["SCS_SKU"]).join(pd.DataFrame(filter_df["ComponentGroup"]).join(filter_df["Component_SCS"]))
    #Eliminar registros duplicados del DataFrame resultante
    final_df=final_df.drop_duplicates()


    #Escribir el DataFrame resultante en un archivo nuevo de excel
    with pd.ExcelWriter("Duplicated_AVs.xlsx", engine='xlsxwriter') as writer:
        #SKuAcc_df.to_excel(writer, index=False, sheet_name='SCS_BOM')
        #MS4_df.to_excel(writer, index=False, sheet_name='MS4_BOM')
        #lookup_df.to_excel(writer, index=False, sheet_name='LOOKUP')
        #filter_df.to_excel(writer, index=False, sheet_name='FILTER')
        final_df.to_excel(writer, index=False, sheet_name='DUPLICATED_LIST')

    print("Done ;D")

    #Armar todos los DF para escribirlos al final, es poco eficiente modificar el contenido directamente en excel
    #Evitar guardar el excel (Sólo guardar para control)