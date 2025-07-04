def generate_subtotals_generic(df, hierarchy_cols, value_cols):
    """
    Génère un DataFrame avec des sous-totaux hiérarchiques,
    uniquement pour les groupes contenant plusieurs éléments.
    
    :param df: DataFrame source
    :param hierarchy_cols: liste hiérarchique, ex: ["Category", "Sub-Category", "Product"]
    :param value_cols: colonnes numériques à agréger, ex: ["Amount1", "Amount2"]
    :return: DataFrame avec sous-totaux hiérarchiques
    """
    base = df.copy()

    # Supprimer la valeur du dernier niveau si redondant avec le précédent
    if len(hierarchy_cols) >= 2:
        last = hierarchy_cols[-1]
        prev = hierarchy_cols[-2]
        base[last] = base.apply(
            lambda row: "" if row[last] == row[prev] else row[last], axis=1
        )

    all_levels = [base]

    # Construction des sous-totaux intermédiaires (niveaux > 0)
    for level in reversed(range(1, len(hierarchy_cols))):
        group_cols = hierarchy_cols[:level]
        subtotal_label = hierarchy_cols[level]

        # Ne garder que les groupes ayant >1 élément enfant
        children_counts = (
            df.groupby(group_cols)[hierarchy_cols[level]]
            .nunique()
            .reset_index()
        )
        valid_groups = children_counts[children_counts[hierarchy_cols[level]] > 1]
        if valid_groups.empty:
            continue

        # Calculer les sous-totaux pour ces groupes valides
        subtotals = (
            df.groupby(group_cols)[value_cols]
            .sum()
            .reset_index()
            .merge(valid_groups[group_cols], on=group_cols)
        )

        # Créer les colonnes manquantes avec vide
        for col in hierarchy_cols[level:]:
            subtotals[col] = ""
        subtotals[subtotal_label] = "SOUS-TOTAL " + subtotals[subtotal_label]

        # Réordonner
        subtotals = subtotals[hierarchy_cols + value_cols]
        all_levels.append(subtotals)

    # Niveau 0 : total principal par le plus haut niveau
    if len(hierarchy_cols) >= 1:
        level0 = hierarchy_cols[0]
        counts_lvl0 = df[level0].value_counts()
        multi0 = counts_lvl0[counts_lvl0 > 1].index.tolist()

        lvl0_total = df.groupby([level0])[value_cols].sum().reset_index()
        lvl0_total = lvl0_total[lvl0_total[level0].isin(multi0)]

        for col in hierarchy_cols[1:]:
            lvl0_total[col] = ""
        lvl0_total[level0] = "SOUS-TOTAL " + lvl0_total[level0]
        lvl0_total = lvl0_total[hierarchy_cols + value_cols]
        all_levels.append(lvl0_total)

    # Fusion de toutes les lignes
    result = pd.concat(all_levels, ignore_index=True)
    result = result.sort_values(by=hierarchy_cols, ignore_index=True)

    return result

















import pandas as pd

def generate_subtotals_generic(df, hierarchy_cols, value_cols):
    """
    Génère un DataFrame avec des sous-totaux hiérarchiques dynamiques.
    
    :param df: DataFrame source
    :param hierarchy_cols: liste ordonnée des colonnes hiérarchiques (ex : ["Cat", "SubCat", "Product"])
    :param value_cols: liste des colonnes à agréger (ex : ["Amount1", "Amount2"])
    :return: DataFrame avec sous-totaux insérés
    """
    base = df.copy()

    # Masquer dernier niveau s'il est égal au précédent
    if len(hierarchy_cols) >= 2:
        last = hierarchy_cols[-1]
        prev = hierarchy_cols[-2]
        base[last] = base.apply(
            lambda row: "" if row[last] == row[prev] else row[last], axis=1
        )

    all_levels = [base]

    # Boucle sur les niveaux hiérarchiques (du plus profond vers le plus haut)
    for level in reversed(range(1, len(hierarchy_cols))):
        group_cols = hierarchy_cols[:level]
        group_name = hierarchy_cols[level]

        # Vérifier si plus d'un élément par parent (sinon pas de sous-total)
        counts = df.groupby(group_cols)[hierarchy_cols[level]].nunique().reset_index()
        multiple = counts[counts[hierarchy_cols[level]] > 1]
        if multiple.empty:
            continue

        # Calcul des sous-totaux
        subtotal = (
            df.groupby(group_cols)[value_cols]
            .sum()
            .reset_index()
            .merge(multiple[group_cols], on=group_cols)
        )
        for col in hierarchy_cols[level:]:
            subtotal[col] = ""
        subtotal[group_name] = "SOUS-TOTAL " + subtotal[group_name]

        # Réordonner les colonnes
        subtotal = subtotal[hierarchy_cols + value_cols]
        all_levels.append(subtotal)

    # Niveau 0 (totaux globaux)
    if len(hierarchy_cols) >= 1:
        level0 = hierarchy_cols[0]
        level0_counts = df[level0].value_counts()
        multi0 = level0_counts[level0_counts > 1].index.tolist()

        total0 = df.groupby([level0])[value_cols].sum().reset_index()
        total0 = total0[total0[level0].isin(multi0)]
        for col in hierarchy_cols[1:]:
            total0[col] = ""
        total0[hierarchy_cols[0]] = "SOUS-TOTAL " + total0[hierarchy_cols[0]]
        total0 = total0[hierarchy_cols + value_cols]
        all_levels.append(total0)

    # Fusion finale
    result = pd.concat(all_levels, ignore_index=True)
    result = result.sort_values(by=hierarchy_cols, ignore_index=True)

    return result






def generate_subtotals_inplace(df, value_column="Amount"):
    # Copie des lignes d'origine
    base_rows = df.copy()

    # Niveau 2 : sous-category
    sub_counts = df.groupby(["Category", "Sub-Category"])["Product"].nunique().reset_index()
    multi_products = sub_counts[sub_counts["Product"] > 1][["Category", "Sub-Category"]]

    lvl2 = (
        df.groupby(["Category", "Sub-Category"], as_index=False)[value_column]
        .sum()
        .merge(multi_products, on=["Category", "Sub-Category"])
    )
    lvl2["Product"] = "SOUS-TOTAL " + lvl2["Sub-Category"]
    lvl2 = lvl2[df.columns]  # assurer même ordre des colonnes

    # Niveau 1 : category
    cat_counts = df.groupby("Category")["Sub-Category"].nunique().reset_index()
    multi_subcats = cat_counts[cat_counts["Sub-Category"] > 1]["Category"]

    lvl1 = (
        df.groupby(["Category"], as_index=False)[value_column]
        .sum()
    )
    lvl1 = lvl1[lvl1["Category"].isin(multi_subcats)]
    lvl1["Sub-Category"] = ""
    lvl1["Product"] = "SOUS-TOTAL " + lvl1["Category"]
    lvl1 = lvl1[df.columns]

    # Combinaison complète
    df_final = pd.concat([base_rows, lvl2, lvl1], ignore_index=True)
    df_final = df_final.sort_values(by=["Category", "Sub-Category", "Product"], ignore_index=True)

    return df_final

















import pandas as pd

def generate_subtotals(df):
    # Niveau produit (ligne brute)
    lvl3 = df.copy()

    # Niveau Sous-Category (niveau 2)
    sub_counts = df.groupby(["Category", "Sub-Category"])["Product"].nunique().reset_index()
    multi_products = sub_counts[sub_counts["Product"] > 1][["Category", "Sub-Category"]]

    lvl2 = (
        df.groupby(["Category", "Sub-Category"], as_index=False)["Amount"]
        .sum()
        .merge(multi_products, on=["Category", "Sub-Category"])
    )
    lvl2["Product"] = "SOUS-TOTAL " + lvl2["Sub-Category"]

    # Niveau Category (niveau 1)
    cat_counts = df.groupby("Category")["Sub-Category"].nunique().reset_index()
    multi_subcats = cat_counts[cat_counts["Sub-Category"] > 1]["Category"]

    lvl1 = (
        df.groupby(["Category"], as_index=False)["Amount"]
        .sum()
    )
    lvl1 = lvl1[lvl1["Category"].isin(multi_subcats)]
    lvl1["Sub-Category"] = ""
    lvl1["Product"] = "SOUS-TOTAL " + lvl1["Category"]

    # Regrouper
    full = pd.concat([lvl3, lvl2, lvl1], ignore_index=True)

    # Trier proprement
    full = full.sort_values(by=["Category", "Sub-Category", "Product"], ignore_index=True)

    return full





import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def update_dashboard_sheet(filepath, sheet_name="Dashboard", data_sheet="Data"):
    # Charger le classeur
    wb = load_workbook(filepath)
    if sheet_name not in wb.sheetnames or data_sheet not in wb.sheetnames:
        print(f"Feuille '{sheet_name}' ou '{data_sheet}' introuvable.")
        return

    ws = wb[sheet_name]
    ws_data = wb[data_sheet]

    # Lire les filtres de la feuille Dashboard
    cdr = ws["C2"].value
    law = ws["C3"].value
    date1 = ws["D6"].value
    date2 = ws["E6"].value
    date3 = ws["F6"].value

    # Lire les données sources dans un DataFrame
    data = ws_data.values
    cols = next(data)
    df = pd.DataFrame(data, columns=cols)

    # Filtrer selon CDR et LAW
    df_filtered = df[(df["CDR"] == cdr) & (df["LAW"] == law)]

    # Construire tableau dynamique (comme un TCD)
    pivot = df_filtered.pivot_table(
        index=["Category", "Sub-Category", "Product"],
        columns="DATE_STOCK",
        values="AMOUNT",
        aggfunc="sum"
    ).reset_index()

    # S'assurer que les dates demandées sont présentes
    for d in [date1, date2, date3]:
        if d not in pivot.columns:
            pivot[d] = None

    # Réorganiser le tableau
    final_table = pivot[["Category", "Sub-Category", "Product", date1, date2, date3]]

    # Nettoyer la zone de tableau (ligne 8+)
    for row in ws.iter_rows(min_row=8, max_col=6, max_row=ws.max_row):
        for cell in row:
            cell.value = None

    # Insérer les nouvelles lignes à partir de la ligne 8
    for i, row in enumerate(dataframe_to_rows(final_table, index=False, header=False)):
        for j, val in enumerate(row):
            ws.cell(row=8 + i, column=1 + j, value=val)

    # Sauvegarde
    updated_path = filepath.replace(".xlsx", "_UPDATED.xlsx")
    wb.save(updated_path)
    print(f"✅ Feuille '{sheet_name}' mise à jour et enregistrée sous : {updated_path}")







Private Sub Worksheet_Change(ByVal Target As Range)
    Dim watchCells As Range
    Set watchCells = Union(Me.Range("C2"), Me.Range("C3"), Me.Range("D6"), Me.Range("E6"), Me.Range("F6"))

    If Not Application.Intersect(Target, watchCells) Is Nothing Then
        Application.EnableEvents = False

        If MsgBox("Valeur modifiée sur " & Me.Name & "." & vbCrLf & "Actualiser le tableau de cette feuille ?", _
                  vbYesNo + vbQuestion, "Actualisation") = vbYes Then
            Call UpdateDashboardForSheet(Me)
        End If

        Application.EnableEvents = True
    End If
End Sub





Can I add another submodule to lanch this code wheever thé user change one of thé 5 Cells values Of course. This is an excellent and very practical use case for VBA. The level of dynamic interaction you're describing is where VBA starts to shine and becomes much more efficient than complex formula-based solutions.
The logic is:
- 
User provides two high-level filters (CDR and LAW).

- 
VBA filters the master Pivot Table based on those inputs.

- 
VBA then reads the resulting, smaller set of row labels (Category, Sub-Category, Product) from the filtered pivot.

- 
VBA populates the dashboard table with these dynamic row labels.

- 
Finally, VBA loops through this new table to fill in the data values for the three dates specified by the user.


Here is the complete solution, including the Excel setup and the VBA code.
### **Part 1: The Excel Template Setup (Crucial!)**
Before you write any code, your Dashboard sheet needs to be set up with named cells and a named table. This makes your VBA code clean and easy to read.
**Step 1: Create Named Input Cells**
- 
Choose cells for your user inputs.

- 
Select the cell where the user will type the CDR value (e.g., C2).

- 
Go to the **Name Box** (the small box to the left of the formula bar) and type the name **inp_CDR** and press **Enter**.

- 
Repeat this for the other inputs:
  - 
C3 -> **inp_LAW**

  - 
D6 -> **inp_Date1**

  - 
E6 -> **inp_Date2**

  - 
F6 -> **inp_Date3**




**Step 2: Create the Presentation Table**
- 
Create the headers for your output table starting in row 7, for example.
  - 
A7: **Category**

  - 
B7: **Sub-Category**

  - 
C7: **Product**

  - 
D7: (This header will be linked to the date input) -> formula =inp_Date1

  - 
E7: -> formula =inp_Date2

  - 
F7: -> formula =inp_Date3



- 
Select all these headers (A7:F7).

- 
Press **Ctrl + T** (or go to **Insert > Table**).

- 
Check **"My table has headers"** and click **OK**.

- 
A new **Table Design** tab appears. On the far left, change the **Table Name** to **tbl_Report**.


Your Dashboard is now set up perfectly for VBA to interact with.
### **Part 2: The VBA Code**
Press **Alt + F11** to open the VBA Editor, go to **Insert > Module**, and paste the following code.
Generated vba
      `' ================================================================================= '  Main Subroutine to update the dynamic dashboard ' ================================================================================= Sub UpdateDynamicDashboard()     ' --- Configuration: Set these names to match your workbook ---     Const DATA_SHEET_NAME As String = "Data"     Const PIVOT_TABLE_NAME As String = "TCD_Data"     Const REPORT_TABLE_NAME As String = "tbl_Report"          ' --- Variable Declaration ---     Dim wb As Workbook     Dim dashboardSheet As Worksheet     Dim dataSheet As Worksheet     Dim pt As PivotTable     Dim reportTable As ListObject          ' User Inputs     Dim cdrInput As String, lawInput As String     Dim date1 As Date, date2 As Date, date3 As Date          ' Pivot Fields for filtering     Dim ptFieldCDR As PivotField, ptFieldLAW As PivotField          ' --- Error Handling ---     On Error GoTo ErrorHandler          ' --- Optimization: Make the macro run much faster ---     Application.ScreenUpdating = False     Application.EnableEvents = False     Application.Calculation = xlCalculationManual          ' --- 1. Get Handles to all our objects ---     Set wb = ThisWorkbook     Set dashboardSheet = wb.ActiveSheet ' Assumes the user is on the dashboard     Set dataSheet = wb.Sheets(DATA_SHEET_NAME)     Set pt = dataSheet.PivotTables(PIVOT_TABLE_NAME)     Set reportTable = dashboardSheet.ListObjects(REPORT_TABLE_NAME)          ' --- 2. Read all user inputs from the named cells ---     cdrInput = dashboardSheet.Range("inp_CDR").Value     lawInput = dashboardSheet.Range("inp_LAW").Value     date1 = dashboardSheet.Range("inp_Date1").Value     date2 = dashboardSheet.Range("inp_Date2").Value     date3 = dashboardSheet.Range("inp_Date3").Value          ' --- 3. Filter the Master Pivot Table based on user inputs ---     Debug.Print "Filtering Pivot for CDR='" & cdrInput & "' and LAW='" & lawInput & "'"          ' Get handles to the pivot fields we need to filter     Set ptFieldCDR = pt.PivotFields("CDR") ' Change "CDR" to your actual field name     Set ptFieldLAW = pt.PivotFields("LAW") ' Change "LAW" to your actual field name          ' Clear old filters to ensure a clean state     pt.ClearAllFilters          ' Apply the new filters from user input     ptFieldCDR.CurrentPage = cdrInput     ptFieldLAW.CurrentPage = lawInput     pt.PivotCache.Refresh ' Ensure pivot cache reflects the filter change          ' --- 4. Populate the Report Table with the new, filtered row labels ---     Debug.Print "Populating report table rows..."          ' Clear any old data from the report table     If reportTable.ListRows.Count > 0 Then         reportTable.DataBodyRange.Delete     End If          ' Loop through the visible items of the filtered pivot table     Dim newRow As ListRow     Dim catItem As PivotItem, subCatItem As PivotItem, prodItem As PivotItem          ' This nested loop structure correctly extracts the hierarchical row data     For Each catItem In pt.PivotFields("Category").VisibleItems         For Each subCatItem In pt.PivotFields("Sub-Category").VisibleItems             For Each prodItem In pt.PivotFields("Product").VisibleItems                 ' Use a quick GetPivotData check to see if this combination actually exists                 On Error Resume Next                 Dim checkValue As Variant                 checkValue = pt.GetPivotData("Somme de AMOUNT", "Category", catItem.Name, "Sub-Category", subCatItem.Name, "Product", prodItem.Name)                 If Err.Number = 0 Then ' If no error, the combination is valid                     Set newRow = reportTable.ListRows.Add                     newRow.Range(1).Value = catItem.Name                     newRow.Range(2).Value = subCatItem.Name                     newRow.Range(3).Value = prodItem.Name                 End If                 On Error GoTo ErrorHandler ' Reset error handling             Next prodItem         Next subCatItem     Next catItem          ' --- 5. Fill in the data values for the three dates ---     Debug.Print "Filling in data values for the dates..."          Dim r As ListRow     Dim pivotValue As Variant     For Each r In reportTable.ListRows         ' Get data for Date 1         On Error Resume Next         pivotValue = pt.GetPivotData("Somme de AMOUNT", "Category", r.Range(1).Value, "Sub-Category", r.Range(2).Value, "Product", r.Range(3).Value, "DATE_STOCK", date1)         r.Range(4).Value = IIf(IsError(pivotValue), "", pivotValue)         Erase pivotValue                  ' Get data for Date 2         pivotValue = pt.GetPivotData("Somme de AMOUNT", "Category", r.Range(1).Value, "Sub-Category", r.Range(2).Value, "Product", r.Range(3).Value, "DATE_STOCK", date2)         r.Range(5).Value = IIf(IsError(pivotValue), "", pivotValue)         Erase pivotValue                  ' Get data for Date 3         pivotValue = pt.GetPivotData("Somme de AMOUNT", "Category", r.Range(1).Value, "Sub-Category", r.Range(2).Value, "Product", r.Range(3).Value, "DATE_STOCK", date3)         r.Range(6).Value = IIf(IsError(pivotValue), "", pivotValue)         Erase pivotValue                  On Error GoTo ErrorHandler     Next r          MsgBox "Dashboard has been updated!", vbInformation  Cleanup:     ' --- Restore Excel's normal settings ---     Application.Calculation = xlCalculationAutomatic     Application.EnableEvents = True     Application.ScreenUpdating = True     Exit Sub  ErrorHandler:     MsgBox "An error occurred in the VBA code: " & vbCrLf & Err.Description, vbCritical     Resume Cleanup End Sub`     
**Important:** You must change the field names in quotes (like "CDR", "LAW", "Category", "Sub-Category", "Product", "DATE_STOCK", and "Somme de AMOUNT") to **exactly match** the names in your Pivot Table.
### **Part 3: How to Run the Code**
- 
**Add a Button:** On your Dashboard sheet, go to **Insert > Shapes** and draw a button. Type "Refresh Report" on it.

- 
**Assign the Macro:** **Right-click** on the button shape and choose **Assign Macro...**.

- 
Select **UpdateDynamicDashboard** from the list and click **OK**.


Now, your user can simply type their CDR and LAW values, choose their three dates, and click the "Refresh Report" button. The VBA code will run and build the entire custom report for them instantly.























Yes, absolutely. This is an excellent idea and a very common and professional way to solve this exact problem.

You are correct to identify that the Oracle database connection is the most fragile part of the system. By moving this responsibility from Excel into a dedicated Python script, you make the entire process more robust, portable, and easier to debug.

Let's go back to that first, simpler architecture. It's a fantastic choice.

The New, More Reliable Architecture

Here is the new workflow. It's simpler and more reliable because the "hard part" (talking to the database) is handled entirely by Python, which is excellent at it.

Step 1 (New Python Script): A Python script connects directly to the Oracle database, runs the SQL query, and saves the result as a clean source_data.csv file.

Step 2 (Excel/Power Query): Your Excel template's Power Query connection now points to that simple source_data.csv file, not to the Oracle database.

Step 3 (Existing Python Script): Your main Python script (main.py and the GUI) works exactly as before, but when it tells Excel to "Refresh", Excel is just refreshing from a local file, which is instant and never fails.

This is a great design because you are moving the complexity into the tool best suited to handle it (Python).

Step-by-Step Guide to Make the Change

Let's modify our project. This is very straightforward.

Part 1: Create the New Python Script for Database Connection

This script will replace the need for an Oracle Client for Excel. You will, however, need a Python library to talk to Oracle. The best one is oracledb.

Step 1: Install the Python-Oracle Library
Open your command prompt or terminal and run:

Generated bash
pip install oracledb


Step 2: Create the connect_and_export.py Script
In your Report_Project folder, create a new Python file named connect_and_export.py. Paste the following code into it.

IMPORTANT: You must fill in your actual database credentials. For security, it's best not to write your password directly in the code. This example shows how to get it from environment variables, which is safer.

Generated python
# connect_and_export.py
import oracledb
import pandas as pd
import os

# --- Configuration: Fill in your Oracle Database details here ---
# For better security, load credentials from environment variables or a secure vault.
# To set an environment variable (Windows example): setx DB_PASSWORD "your_password"
# Then restart your terminal.
DB_USER = "your_username"
DB_PASSWORD = os.getenv("ORACLE_DB_PASSWORD", "your_password_if_not_set")
DB_HOST = "oracledb.mycompany.com"
DB_PORT = 1521
DB_SERVICE_NAME = "salespdb.mycompany.com"

# The output file that Power Query will read from.
OUTPUT_CSV_FILE = "source_data.csv"

# --- Main Logic ---
print("Attempting to connect to the Oracle database...")

# Construct the DSN (Data Source Name)
dsn = f"{DB_HOST}:{DB_PORT}/{DB_SERVICE_NAME}"

try:
    # Establish the connection
    with oracledb.connect(user=DB_USER, password=DB_PASSWORD, dsn=dsn) as connection:
        print("Successfully connected to Oracle Database!")
        
        # Define the SQL query to fetch all necessary data.
        # It's better to fetch a wider range of data here and let Excel filter it.
        sql_query = """
        SELECT
            SALE_DATE,
            REGION,
            PRODUCT,
            SALES
        FROM
            YOUR_SALES_TABLE
        WHERE
            SALE_DATE >= TO_DATE('2023-01-01', 'YYYY-MM-DD') -- Or a wider date range
        """
        
        print("Executing SQL query...")
        # Use pandas to read the SQL query results directly into a DataFrame
        df = pd.read_sql(sql_query, connection)
        
        print(f"Query successful. Found {len(df)} rows of data.")
        
        # Save the DataFrame to a CSV file
        df.to_csv(OUTPUT_CSV_FILE, index=False)
        
        print(f"Data has been successfully saved to '{OUTPUT_CSV_FILE}'")

except Exception as e:
    print(f"--- AN ERROR OCCURRED ---")
    print(f"Failed to connect to the database or run the query: {e}")
    # Exit with an error code so other scripts know something went wrong
    exit(1)
IGNORE_WHEN_COPYING_START
content_copy
download
Use code with caution.
Python
IGNORE_WHEN_COPYING_END
Part 2: Modify Your Excel Template

Now, we just need to tell your Report_Template.xlsx to read from the new CSV file instead of the database.

Open Report_Template.xlsx.

Go to the Data tab and click Queries & Connections. A pane will open on the right.

You should see your main data query (e.g., SalesData). Right-click on it and choose Edit.

The Power Query Editor will open. On the right side, in the "Applied Steps" pane, find the first step, which is named Source.

Click the little gear icon next to the "Source" step.

This will open a dialog box. It's currently set to Oracle. We need to change this completely.

Delete the "Source" step by clicking the 'X' next to it.

Now, with the query selected, click New Source > File > Text/CSV.

Find and select the source_data.csv file you created with the new Python script.

Power Query will automatically create a new "Source" and "Change Type" step.

That's it! The rest of your query steps (like filtering) should now apply to the data from the CSV file. You can delete any old steps that are no longer needed.

Click Close & Load.

Save and close your Report_Template.xlsx file.

Part 3: Update the Final Workflow

Your process is now a simple two-step process for the user:

New Workflow:

Run connect_and_export.py: This script connects to Oracle and creates the fresh source_data.csv.

Run gui_launcher.py: This opens the pop-up window. The user enters their parameters and clicks "Generate". The main.py script then opens the Excel template and tells it to refresh. Excel will now refresh from the local CSV file, which is fast and reliable.

Optional - Advanced Improvement:
You can even modify your gui_launcher.py script to run both Python scripts in order, so the user only has to do one thing!

Generated python
# In gui_launcher.py, inside the save_and_run_report function

# ...
status_label.config(text="Status: Connecting to database...", fg="blue")
window.update_idletasks()

try:
    # --- STEP 1: Run the database export script ---
    subprocess.run([sys.executable, 'connect_and_export.py'], check=True)
    
    status_label.config(text="Status: Generating Excel report...", fg="blue")
    window.update_idletasks()
    
    # --- STEP 2: Run the main Excel generation script ---
    subprocess.run([sys.executable, 'main.py'], check=True)
    
    status_label.config(text="Status: Report generation successful!", fg="green")
    messagebox.showinfo("Success", "The report has been generated successfully.")

except subprocess.CalledProcessError as e:
    status_label.config(text="Status: Error! Check the console window.", fg="red")
    messagebox.showerror("Error", "An error occurred during the process.")
IGNORE_WHEN_COPYING_START
content_copy
download
Use code with caution.
Python
IGNORE_WHEN_COPYING_END

By making this change, you have created a much more stable and professional automation system. You have successfully isolated the most difficult part of the process.







Of course. Let's build this entire project from scratch using simple, step-by-step language. We will go from an empty folder to a fully working application.

The Goal: What We Are Building

We are creating a system that automatically generates a special Excel report.

A user will open a simple pop-up window (a GUI).

They will enter a start and end date, and list the regions they want a report for (e.g., "North America, Europe").

They click a "Generate" button.

The program automatically connects to an Oracle database, gets the data for that date range, and creates a beautiful, interactive Excel file.

The final Excel file will have a separate sheet for each region requested.

Part 1: Setting Up Your Project

First, let's get our folder and files organized.

Step 1: Create Your Project Folder
Create a new folder on your computer. Let's call it Report_Project. Inside this folder, we will create all our files.

Step 2: Check Your Tools (Prerequisites)
Before we start, make sure you have:

Python: You need Python installed on your computer.

Excel: You need a modern version of Excel (2016 or newer).

Oracle Client: This is special software that lets Excel and Python talk to your Oracle database. Crucially, if you have 64-bit Excel, you need the 64-bit Oracle Client. If you have 32-bit Excel, you need the 32-bit client. This is the most common place people get stuck!

Python Libraries: Open your command prompt (cmd) or terminal and install two libraries with this command:

Generated bash
pip install xlwings configparser

Part 2: Creating the Excel Template (The "Brain")

This is the most important part. We will create a special Excel file that acts as a blueprint for our reports.

Step 2A: Create the Excel File and a "Parameters" Sheet

Open a new, blank Excel workbook.

Save it inside your Report_Project folder as Report_Template.xlsx.

Rename the first sheet to _Parameters.

In cell A1, type ParamName. In cell B1, type ParamValue.

In cell A2, type MinDate. In cell B2, type a placeholder date like 2023-01-01.

In cell A3, type MaxDate. In cell B3, type a placeholder date like 2023-12-31.

Click anywhere on this data (e.g., cell A1), then go to the Insert tab and click Table. Make sure "My table has headers" is checked and click OK.

A new Table Design tab will appear. On the far left, change the Table Name to tbl_Params and press Enter.

Step 2B: Connect Excel to the Oracle Database (with Power Query)

Go to the Data tab in Excel.

Click Get Data > From Database > From Oracle Database.

Enter your Oracle server details and click OK.

The "Navigator" window will appear, showing you all the tables in the database. Find and select your main sales table (e.g., YOUR_SALES_TABLE). Do not click Load yet!

Click the Transform Data button at the bottom. This opens the Power Query Editor.

Now, let's add our date filters. Find your date column (e.g., SALE_DATE). Click the small filter arrow on that column's header.

Go to Date Filters > Between....

In the "Filter Rows" window:

Set the first condition to is after or equal to. In the box next to it, do not pick a date. Instead, type =pMinDate.

Set the second condition to is before or equal to. In the box next to it, type =pMaxDate.

Click OK. (We haven't created pMinDate and pMaxDate yet, so you might see an error. Ignore it for now!)

Now, we'll create those pMinDate and pMaxDate queries. In the Power Query editor, go to Home > New Source > Other Sources > Blank Query.

A new query appears. Rename it to pMinDate.

Click on Advanced Editor and paste this exact code:

Generated m
let Source = Excel.CurrentWorkbook(){[Name="tbl_Params"]}[Content], Value = Source{[ParamName="MinDate"]}[ParamValue] in Value
IGNORE_WHEN_COPYING_START
content_copy
download
Use code with caution.
M
IGNORE_WHEN_COPYING_END

Repeat steps 9-11 to create another blank query named pMaxDate, but use this code instead:

Generated m
let Source = Excel.CurrentWorkbook(){[Name="tbl_Params"]}[Content], Value = Source{[ParamName="MaxDate"]}[ParamValue] in Value
IGNORE_WHEN_COPYING_START
content_copy
download
Use code with caution.
M
IGNORE_WHEN_COPYING_END

Click back on your main sales data query. The error should be gone!

Finally, click the Close & Load dropdown button and choose Close & Load To....

Select Only Create Connection and check the box Add this data to the Data Model. Click OK.

Step 2C: Create the Master Pivot Table

Create a new sheet and name it Data.

Go to Insert > PivotTable.

Choose From Data Model. Click OK.

A new Pivot Table will appear. Go to the PivotTable Analyze tab and name it TCD_Data.

Set it up like this:

Filters: Region (or whatever you use for categories)

Rows: Product

Columns: Your date column (e.g., SALE_DATE)

Values: Sales (or your number column)

Step 2D: Create the Interactive Dashboard Sheet

Create one last sheet and name it Dashboard. This is what the user will see.

Set up three cells for the user to type dates, for example: B1, C1, D1.

Create the frame of your report table below. For example, in A5 put the Product header, and starting in A6, list your products.

In the column headers (e.g., B5, C5, D5), link them to the input cells. For example, the formula in cell B5 is just =B1.

Now for the magic formula. In the first data cell of your report (e.g., B6), write this GETPIVOTDATA formula. You must change the parts in quotes to match your actual field names!

Generated excel
=IFERROR(GETPIVOTDATA("Sum of Sales",Data!$A$3,"Product",$A6,"SALE_DATE",B$5),"")
IGNORE_WHEN_COPYING_START
content_copy
download
Use code with caution.
Excel
IGNORE_WHEN_COPYING_END

"Sum of Sales": The name of your value field in the pivot table.

Data!$A$3: Any cell inside your pivot table on the Data sheet.

"Product", $A6: Matches the row by looking at the Product name in column A.

"SALE_DATE", B$5: Matches the column by looking at the date in row 5.

IFERROR(..., ""): If no data is found, it will show a blank cell instead of an error.

Drag this formula across all the rows and columns of your presentation table.

Save your Report_Template.xlsx file and close it.

Part 3: Writing the Python Scripts

Now we'll create the three text files that contain our code.

Step 3A: The Configuration File (config.ini)
This file holds all the settings so you don't have to change the code later.

In your Report_Project folder, create a new text file named config.ini.

Paste the following content into it. You can change these values later.

Generated ini
[Settings]
# What regions/sheets should the script create? Separate with a comma.
sheets_to_create = North America, Europe, Asia

# The default dates that will appear in the pop-up window.
default_date_1 = 2024-01-01
default_date_2 = 2024-03-31

[Paths]
template_file = Report_Template.xlsx
output_directory = ./Generated_Reports/

[Excel]
# Names of things inside your Excel template. Must match exactly!
template_sheet_name = Dashboard
data_sheet_name = Data
data_pivot_name = TCD_Data
parameters_sheet_name = _Parameters
filter_field_name = Region
IGNORE_WHEN_COPYING_START
content_copy
download
Use code with caution.
Ini
IGNORE_WHEN_COPYING_END

Step 3B: The Pop-Up Window (gui_launcher.py)
This script creates the friendly window for the user.

Create a new Python file named gui_launcher.py.

Paste this code into it:

Generated python
import tkinter as tk
from tkinter import messagebox
import configparser
import subprocess
import sys
from datetime import datetime

CONFIG_FILE = 'config.ini'
MAIN_SCRIPT = 'main.py'

def load_config_to_gui():
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE)
    entry_start_date.insert(0, config.get('Settings', 'default_date_1'))
    entry_end_date.insert(0, config.get('Settings', 'default_date_2'))
    entry_sheets.insert(0, config.get('Settings', 'sheets_to_create'))

def save_and_run_report():
    # 1. Read values from the GUI
    start_date_str = entry_start_date.get()
    end_date_str = entry_end_date.get()
    sheets_str = entry_sheets.get()

    # 2. Check if dates are in the correct format
    try:
        datetime.strptime(start_date_str, '%Y-%m-%d')
        datetime.strptime(end_date_str, '%Y-%m-%d')
    except ValueError:
        messagebox.showerror("Invalid Date", "Please enter dates in YYYY-MM-DD format.")
        return

    # 3. Save the user's choices back to the config file
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE)
    config['Settings']['default_date_1'] = start_date_str
    config['Settings']['default_date_2'] = end_date_str
    config['Settings']['sheets_to_create'] = sheets_str
    with open(CONFIG_FILE, 'w') as configfile:
        config.write(configfile)

    # 4. Tell the user we are running the report
    status_label.config(text="Status: Generating report... Please wait.", fg="blue")
    window.update_idletasks() # Force the GUI to update immediately

    # 5. Run the main script that does all the heavy lifting
    try:
        # Use sys.executable to make sure we run with the same Python
        subprocess.run([sys.executable, MAIN_SCRIPT], check=True)
        status_label.config(text="Status: Report generation successful!", fg="green")
        messagebox.showinfo("Success", "The report has been generated successfully.")
    except subprocess.CalledProcessError as e:
        status_label.config(text="Status: Error! Check the console window for details.", fg="red")
        messagebox.showerror("Error", "An error occurred during report generation.")

# --- This part builds the actual window ---
window = tk.Tk()
window.title("Report Generator")
frame = tk.Frame(window, padx=10, pady=10)
frame.pack(expand=True)

tk.Label(frame, text="Start Date for Query (YYYY-MM-DD):").grid(row=0, column=0, sticky="w", pady=2)
entry_start_date = tk.Entry(frame, width=40)
entry_start_date.grid(row=0, column=1, pady=2)

tk.Label(frame, text="End Date for Query (YYYY-MM-DD):").grid(row=1, column=0, sticky="w", pady=2)
entry_end_date = tk.Entry(frame, width=40)
entry_end_date.grid(row=1, column=1, pady=2)

tk.Label(frame, text="Sheets to Create (comma-separated):").grid(row=2, column=0, sticky="w", pady=2)
entry_sheets = tk.Entry(frame, width=40)
entry_sheets.grid(row=2, column=1, pady=2)

run_button = tk.Button(frame, text="Generate Report", command=save_and_run_report)
run_button.grid(row=3, column=0, columnspan=2, pady=20)

status_label = tk.Label(frame, text="Status: Ready", fg="gray")
status_label.grid(row=4, column=0, columnspan=2, pady=5)

load_config_to_gui()
window.mainloop()
IGNORE_WHEN_COPYING_START
content_copy
download
Use code with caution.
Python
IGNORE_WHEN_COPYING_END

Step 3C: The Main Script (main.py)
This is the engine of our project. It reads the settings, controls Excel, and builds the final report.

Create a new Python file named main.py.

Paste this code into it:

Generated python
import xlwings as xw
import time
import os
import configparser
from datetime import datetime

def main():
    # --- Part 1: Read all settings from the config file ---
    config = configparser.ConfigParser()
    config.read('config.ini')

    settings = config['Settings']
    paths = config['Paths']
    excel_config = config['Excel']

    min_date = settings.get('default_date_1')
    max_date = settings.get('default_date_2')
    sheets_to_create = [s.strip() for s in settings.get('sheets_to_create').split(',')]
    
    template_path = os.path.abspath(paths.get('template_file'))
    output_dir = paths.get('output_directory')
    os.makedirs(output_dir, exist_ok=True) # Create the output folder if it doesn't exist

    # --- Part 2: Open Excel and tell it to get the data ---
    # The "with" statement makes sure Excel closes properly, even if there's an error
    with xw.App(visible=False) as app:
        try:
            wb = app.books.open(template_path)

            print(f"Telling Excel to get data between {min_date} and {max_date}...")
            
            # Write the dates from our GUI into the _Parameters sheet in Excel
            params_sheet = wb.sheets[excel_config.get('parameters_sheet_name')]
            params_sheet.range("B2").value = min_date
            params_sheet.range("B3").value = max_date

            # Tell Excel to refresh all its data connections (this runs Power Query)
            wb.api.RefreshAll()
            
            # IMPORTANT: We have to wait for the data to finish loading
            while wb.api.Application.BackgroundQueryStatus > 0:
                print("Waiting for data to load from Oracle...")
                time.sleep(2)
            print("Data has been loaded successfully!")

            # --- Part 3: Build the final report by copying sheets ---
            dashboard_template = wb.sheets[excel_config.get('template_sheet_name')]
            master_pivot = wb.sheets[excel_config.get('data_sheet_name')].pivot_tables[excel_config.get('data_pivot_name')]
            filter_field = master_pivot.api.PivotFields(excel_config.get('filter_field_name'))

            # Get a list of all valid regions that exist in the loaded data
            valid_pivot_items = [item.Name for item in filter_field.PivotItems()]

            print(f"Now creating a sheet for each requested region...")
            for sheet_name in sheets_to_create:
                if sheet_name in valid_pivot_items:
                    # Filter the master pivot table to just this one region
                    filter_field.CurrentPage = sheet_name
                    # Copy the beautiful dashboard template
                    dashboard_template.copy(after=wb.sheets[-1], name=sheet_name)
                    print(f"  - Created sheet for '{sheet_name}'")
                else:
                    print(f"  - WARNING: Could not create sheet for '{sheet_name}'. It was not found in the data for this date range.")

            # --- Part 4: Clean up and save the final report ---
            # Hide the template sheets so the user doesn't see them
            for sheet_name in [excel_config.get('parameters_sheet_name'), excel_config.get('data_sheet_name'), excel_config.get('template_sheet_name')]:
                if sheet_name in [s.name for s in wb.sheets]:
                    wb.sheets[sheet_name].visible = False
            
            # Save the final file with a unique name
            output_filename = f"Interactive_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsb"
            output_path = os.path.join(output_dir, output_filename)
            wb.save(output_path)
            
            print(f"\nSUCCESS! Your report is ready.")
            print(f"It has been saved here: {output_path}")

        except Exception as e:
            print(f"\n--- AN ERROR OCCURRED! ---")
            print(e)
            # This makes sure the error is passed back to the GUI
            raise

if __name__ == "__main__":
    main()
IGNORE_WHEN_COPYING_START
content_copy
download
Use code with caution.
Python
IGNORE_WHEN_COPYING_END
Part 4: How to Run Your New Application

This is the easy part!

Navigate to your Report_Project folder.

Find the file gui_launcher.py and double-click it.

The pop-up window will appear.

Enter the dates and the comma-separated list of regions you want.

Click the "Generate Report" button.

Wait a few moments. A black command window will be working in the background. When it's done, you'll see a "Success" message.

Look inside the new Generated_Reports folder. Your brand new, interactive Excel file will be there
