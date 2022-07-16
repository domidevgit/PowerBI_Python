#install AMO/ADOMD https://docs.microsoft.com/en-us/analysis-services/client-libraries?view=azure-analysis-services-current

import clr
def _load_assemblies(amo_path=None, adomd_path=None):
    """
    Loads required assemblies, called after function definition.
    Might need to install SSAS client libraries:
    https://docs.microsoft.com/en-us/azure/analysis-services/analysis-services-data-providers

    Parameters
    ----------
    amo_path : str, default None
        The full path to the DLL file of the assembly for AMO. 
        Should end with '**Microsoft.AnalysisServices.Tabular.dll**'
        Example: C:/my/path/to/Microsoft.AnalysisServices.Tabular.dll
        If None, will use the default location on Windows.
    adomd_path : str, default None
        The full path to the DLL file of the assembly for ADOMD. 
        Should end with '**Microsoft.AnalysisServices.AdomdClient.dll**'
        Example: C:/my/path/to/Microsoft.AnalysisServices.AdomdClient.dll
        If None, will use the default location on Windows.
    """
    # Full path of .dll files
    root = Path(r"C:\Windows\Microsoft.NET\assembly\GAC_MSIL")
    # get latest version of libraries if multiple libraries are installed (max func)
    if amo_path is None:
        amo_path = str(
            max((root / "Microsoft.AnalysisServices.Tabular").iterdir())
            / "Microsoft.AnalysisServices.Tabular.dll"
        )
    if adomd_path is None:
        adomd_path = str(
            max((root / "Microsoft.AnalysisServices.AdomdClient").iterdir())
            / "Microsoft.AnalysisServices.AdomdClient.dll"
        )

    # load .Net assemblies
    logger.info("Loading .Net assemblies...")
    clr.AddReference("System")
    clr.AddReference("System.Data")
    clr.AddReference(amo_path)
    clr.AddReference(adomd_path)

    # Only after loaded .Net assemblies
    global System, DataTable, AMO, ADOMD

    import System
    from System.Data import DataTable
    import Microsoft.AnalysisServices.Tabular as AMO
    import Microsoft.AnalysisServices.AdomdClient as ADOMD

    logger.info("Successfully loaded these .Net assemblies: ")
    for a in clr.ListAssemblies(True):
        logger.info(a.split(",")[0])

TOMServer = AMO.Server()
TOMServer.Connect("localhost:60469") # BI Port can be read from Dax Studio!

for item in TOMServer.Databases:
    print("Database: ", item.Name)
    print("Compatibility Level: ", item.CompatibilityLevel) 
    print("Created: ", item.CreatedTimestamp)

PowerBIDatabase = TOMServer.Databases["a374aed2-6e45-4042-ae76-5190667a8e1e"]

# Define measure dataframe
dfMeasures = pd.DataFrame(
    columns=['Table',
             'Name', 
             'Description', 
             'DataType', 
             'DataCategory',
             'Expression',
             'FormatString',
             'DisplayFolder',
             'Implicit',
             'Hidden',
             'ModifiedTime',
             'State'])

# Define column dataframe
dfColumns = pd.DataFrame(
    columns=['Table',
             'Name'])

# Tables
print("Listing tables...")
for table in PowerBIDatabase.Model.Tables:
    print(table.Name)

    # Assign current table by name
    CurrentTable = PowerBIDatabase.Model.Tables.Find(table.Name)

    # print(type(CurrentTable))
    # print(type(CurrentTable.Measures))

    # Measures
    for measure in CurrentTable.Measures:
        new_row = {'Table':table.Name,
                'Name':measure.Name, 
                'Description':measure.Description, 
                'DataType':measure.DataType,
                'DataCategory':measure.DataCategory,
                'Expression':measure.Expression,
                'FormatString':measure.FormatString,
                'DisplayFolder':measure.DisplayFolder,
                'Implicit':measure.IsSimpleMeasure,
                'Hidden':measure.IsHidden,
                'ModifiedTime':measure.ModifiedTime,
                'State':measure.State}
        #print(new_row)
        dfMeasures = dfMeasures.append(new_row, ignore_index=True)

    # Columns
    for column in CurrentTable.Columns:
        new_row = {'Table':table.Name,
                'Name':column.Name}
        #print(column.Name)
        dfColumns = dfColumns.append(new_row, ignore_index=True)

print(dfMeasures)
print(dfColumns)