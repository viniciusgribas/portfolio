---
title: ETL Process - Brazilian National Oil and Gas Agency (ANP) data.
layout: post
post-image: https://images.pexels.com/photos/162568/oil-pump-jack-sunset-clouds-silhouette-162568.jpeg?auto=compress&cs=tinysrgb&dpr=2&h=750&w=1260
description: This project is a simple ETL where there is an Excel automation via VBA and then a data extraction via Python. The programming languages used were PYTHON and VBA EXCEL. The automation extracts data from a pivot table in excel, converts it to spreadsheets and joins them in python. Preserving all the original files, without modifying them permanently.
tags:
- EXCEL
- VBA EXCEL
- Python
- Pandas
- Xlwgins
- Openpyxl 
---
# [Excel VBA Automation and python ETL]

This project is a simple ETL where there is an Excel automation via VBA and then a data extraction via Python. The programming languages used were PYTHON and VBA EXCEL. The automation extracts data from a pivot table in excel, converts it to spreadsheets and joins them in python. Preserving all the original files, without modifying them permanently.

## Description

1) At first, the files are only available in excel `".XLS"` format, under the name [`"vendas-combustíveis-m3.xls"`](https://github.com/viniciusgribas/ANP-PROJECT/tree/main/assets).

2) Within this initial file, there are two pivot tables that are the target. These are:

    - Pivot Table 1 ) "Vendas, pelas distribuidoras, dos derivados combustíveis de petróleo por Unidade da Federação e produto - 2000-2020 (m3)"
    
    - Pivot Table 2 ) "Vendas, pelas distribuidoras, de óleo diesel por tipo e Unidade da Federação - 2013-2020 (m3)"

3) This data, presented by the pivot tables, does not have its data source easily accessible in another spreadsheet. Also, the data is not available through the Excel shortcut: PivotTableTools>Analyze>Change Data Source. This shows the need to extract them using Excel's own `VBA` programming language. The advantage of extracting them this way is not only the reduced time for processes that could be long, but the possibility of applying them via `python`, through the *[`xlwings library`](https://www.xlwings.org/)*.

     - The worksheet, once opened, has by default only one sheet, called "plan1".
     
     - The macros created in VBA are available in the folder [`"\ANP-PROJECT\Codigos_VBA"`](https://github.com/viniciusgribas/ANP-PROJECT/tree/main/Codigos_VBA).
     - To extract this data, 4 macros were created in VBA. These are presented and described in **PART 2 EXCEL**

4) Once all the *VBA - MACROS* have been created, they can be called by `python` and applied there via *[`xlwings library`](https://www.xlwings.org/)*.

5) After applying the Macros on python, the end products of the extraction are two files in `"CSV-UTF8"`:

    - [`PlanConsolidada1.CSV`](https://github.com/viniciusgribas/ANP-PROJECT/tree/main/assets)
       
    - [`PlanConsolidada2.CSV`](https://github.com/viniciusgribas/ANP-PROJECT/tree/main/assets)

6) These files were managed via the *[`Pandas Library`](https://pandas.pydata.org/)* from python. Having the descriptive in **PART 3 PYTHON**

7) Finally, the final product of this project is two files in `"CSV-UTF8"` available in the folder [`"\ANP-PROJECT\Planilhas Finais"`](https://github.com/viniciusgribas/ANP-PROJECT/tree/main/Planilhas%20Finais), according to the following table:

| Column     | Type      |
|------------|-----------|
| year_month | date      |
| uf         | string    |
| product    | string    |
| unit       | string    |
| volume     | double    |
| created_at | timestamp |

   - [`Sales_Of_Diesel_By_UF_And_Type.CSV`](https://github.com/viniciusgribas/ANP-PROJECT/tree/main/Planilhas%20Finais)

   - [`Sales_Of_Oil_Derivative_Fuels_By_UF_And_Product.CSV`](https://github.com/viniciusgribas/ANP-PROJECT/tree/main/Planilhas%20Finais)


---
# GITHUB: [ANP-PROJECT - [EN]](https://github.com/viniciusgribas/ANP-PROJECT)
---