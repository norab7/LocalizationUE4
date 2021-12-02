using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace TranslationEditor
{
    public static class ExcelSerializer
    {
        // Service Data Row Text
        private const string serviceData = "--== !!! DO NOT TRANSLATE THE TEXT BELOW !!! == SERVICE DATA ==--";

        // Some values so that they are ALWAYS the same spelling regardless how they are used (damn typo bugs) and that all words looked weird when i did this
        // Sheets
        private const string IMPORT = "Import";
        private const string CONTROL = "Control";
        private const string CULTURE = "Culture";

        // Column Headers
        private const string ID = "ID";
        private const string KEY = "Key";
        private const string CONTEXT = "Context";
        private const string SOURCE = "Source";
        private const string TRANSLATION = "Translation";
        private const string DONE = "Done";
        private const string COMMENT = "Comment";
        private const string NAMESPACE = "Namespace";
        private const string PATH = "Path";

        private static string MakeName(string Namespace, string Key)
        {
            return Namespace + ',' + Key;
        }

        private static string GetKey(string ExcelName)
        {
            string[] result = ExcelName.Split(',');
            if (result.Length != 2)
                throw new FormatException("Invalid ExcelName: " + ExcelName + "!");
            return result[1];
        }

        private static string SafeMultilineText(string Value)
        {
            // replace \n to \r\n
            return Regex.Replace(Value, "(?<!\r)\n", "\r\n");
        }

        // Create single worksheet based on sheetname and columns
        private static ExcelWorksheet CreateSingleWorksheet(ref ExcelPackage package, string sheetName, List<string> columns)
        {
            var worksheet = package.Workbook.Worksheets.Add(sheetName);
            if (columns.Count > 0)
            {
                for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
                {
                    worksheet.Cells[1, columnIndex + 1].Value = columns[columnIndex];
                }
            }
            return worksheet;
        }

        // Create a list of worksheets based on a list of sheetnames and columns
        private static List<ExcelWorksheet> CreateMultipleWorksheets(ref ExcelPackage package, List<string> cultureNames, List<string> columns)
        {
            var cultureSheets = new List<ExcelWorksheet>();
            for (var cultureIndex = 0; cultureIndex < cultureNames.Count; cultureIndex++)
            {
                cultureSheets.Add(CreateSingleWorksheet(ref package, cultureNames[cultureIndex], columns));
            }
            return cultureSheets;
        }

        // Get the cell column index where the cell value is equal to header
        private static int GetColumnIndexFromHeader(ExcelWorksheet Worksheet, string header)
        {
            for (var columnIndex = 1; columnIndex <= Worksheet.Dimension.Columns; columnIndex++)
            {
                if ((string)Worksheet.Cells[1, columnIndex].Value == header)
                {
                    return columnIndex;
                }
            }
            return -1;
        }

        // Set a worksheet row to record data based on column header values
        private static void SetSingleWorksheetRowByColumnData(ref ExcelPackage package, string sheetName, int rowIndex, List<string> columns, InternalRecord rec)
        {
            foreach (string column in columns)
            {
                int columnIndex = GetColumnIndexFromHeader(package.Workbook.Worksheets[sheetName], column);
                if (columnIndex != -1)
                {
                    package.Workbook.Worksheets[sheetName].Cells[rowIndex, columnIndex].Value = rec[(column == TRANSLATION ? sheetName : column)];
                }
            }
        }

        // Set multiple worksheets rows to record data based on column header values
        private static void SetMultipleWorksheetRowsByColumnData(ref ExcelPackage package, List<string> sheetNames, int rowIndex, List<string> columns, InternalRecord rec)
        {
            for (var worksheetIndex = 0; worksheetIndex < sheetNames.Count; worksheetIndex++)
            {
                SetSingleWorksheetRowByColumnData(ref package, package.Workbook.Worksheets[sheetNames[worksheetIndex]].Name, rowIndex, columns, rec);
            }
        }

        private static void FormatSheetHeader(ExcelWorksheet sheet)
        {
            // Some parts of formatting are so much easier to just do manually after document export
            if (sheet.Name != "Import")
            {
                // Appearance
                sheet.Row(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Row(1).Style.Fill.BackgroundColor.SetColor(Color.Orange);
                sheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                // sheet.Row(1).Style.Font.Bold = true;

                // Freeze and Filter (Easier to do post upload on the actual googlesheet
                //sheet.View.FreezePanes(2, 1);
                //sheet.Cells[sheet.Dimension.Address].AutoFilter = true;
            }

            // AutoWidth
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
        }

        public static InternalFormat Import_New(string fileName)
        {
            // Empty Data Object
            var data = new InternalFormat();
            data.Cultures = new List<string>(); // hard add 'en'
            data.NativeCulture = "en";
            //data.Cultures.Add("en");

            // Prime sheet names
            var primeSheetNames = new List<string>() { "Import", "Control" };

            // Excel Object Data
            var fileInfo = new FileInfo(fileName);
            var Package = new ExcelPackage(fileInfo);
            var worksheetMap = new Dictionary<string, ExcelWorksheet>();

            // Add each worksheet to map with culture key and add key to data cultures
            foreach (var worksheet in Package.Workbook.Worksheets)
            {
                if (!primeSheetNames.Contains(worksheet.Name))
                {
                    var cultureName = worksheet.Name;
                    worksheetMap.Add(cultureName, worksheet);
                    data.Cultures.Add(cultureName);
                }
            }

            // Create empty records list and get importsheet data
            var ImportSheet = Package.Workbook.Worksheets["Import"];
            var ImportRowCount = ImportSheet.Dimension.Rows;
            List<InternalRecord> records = new List<InternalRecord>(ImportRowCount / 2);

            // Iterate each row in ImportSheet
            for (var row = 1; row <= ImportRowCount; row++)                             // Row = 1 as it uses count NOT indices
            {
                var ImportKey = ImportSheet.Cells[row, 1].Text;

                // Create single record and populate with values from import sheet
                InternalRecord record = new InternalRecord();
                record.Key = ImportKey;
                record.Translations = new List<InternalText>(ImportRowCount);
                record.Source = SafeMultilineText(ImportSheet.Cells[row, 2].Text);      // Source is the base english text for translation
                record.Path = ImportSheet.Cells[row, 4].Text;

                // Add native 'en' translation from ImportSheet
                InternalText nativeTranslation = new InternalText();
                nativeTranslation.Culture = "en";
                nativeTranslation.Text = SafeMultilineText(ImportSheet.Cells[row, 2].Text);
                record.Translations.Add(nativeTranslation);

                // Iterate each translation sheet for key and translation text
                foreach (string culture in worksheetMap.Keys)
                {
                    if(culture == "en") { continue; }                       // bypass 'en' for now because it's the native language

                    bool flag_keyFound = false;

                    var cultureSheet = worksheetMap[culture];
                    var cultureRowCount = cultureSheet.Dimension.Rows;
                    

                    InternalText translation = new InternalText();
                    translation.Culture = culture;

                    var cultureKeyColumn = GetColumnIndexFromHeader(cultureSheet, KEY);
                    var cultureTranslationColumn = GetColumnIndexFromHeader(cultureSheet, TRANSLATION);


                    // Iterate each row in the culturesheet for matching import key value
                    for (var cultureRow = 2; cultureRow <= cultureRowCount; cultureRow++)           // <= count because it's all using count NOT indices
                    {
                        var cultureKey = cultureSheet.Cells[cultureRow, cultureKeyColumn].Text;

                        // If matching import/culture key, add translation text to translation object and break;
                        if (ImportKey == cultureKey)
                        {
                            translation.Text = SafeMultilineText(cultureSheet.Cells[cultureRow, cultureTranslationColumn].Text);
                            flag_keyFound = true;
                            break;
                        }
                    }

                    if (flag_keyFound == false)
                    {
                        throw new FormatException("Culture does not contain translation.\n\nKey: " + ImportKey + "\nCulture: " + culture + "\nTranslation: " + nativeTranslation.Text + "\n\nReturn to google sheets and resync.");
                    }
                    // Add translation object to record
                    record.Translations.Add(translation);
                }
                // Add entire record of all translations to List<InternalRecord> object
                records.Add(record);
            }


            data.Namespaces = new List<InternalNamespace>();
            InternalNamespace lastNS = null;

            // Iterate Import again to reorganise records into Namespaces
            for (var row = 1; row <= ImportRowCount; row++)
            {
                // Get stored record for comparison and appending
                InternalRecord record = records[row - 1];
                var recordKey = record.Key;
                var ImportKey = ImportSheet.Cells[row, 1].Text;

                // Console.WriteLine("Import/Record: " + ImportKey + " / " + recordKey);

                if (ImportKey != recordKey)
                {
                    Console.WriteLine("UnexpectedKey: " + recordKey);
                    throw new FormatException("Unexpected Key: " + recordKey);
                }

                // Is stored namespace null or different to current row namespace
                var ns = ImportSheet.Cells[row, 3].Text;
                if (lastNS == null || lastNS.Name != ns)
                {
                    // Create new namespace and add to data
                    lastNS = new InternalNamespace();
                    lastNS.Name = ns;
                    lastNS.Children = new List<InternalRecord>();
                    data.Namespaces.Add(lastNS);
                    // Console.WriteLine("NewNamespace: " + ns);
                }
                lastNS.Children.Add(record);
            }

            return data;
        }

        public static InternalFormat Import(string FileName)
        {
            var data = new InternalFormat();

            var fileInfo = new FileInfo(FileName);
            var Package = new ExcelPackage(fileInfo);
            var Worksheet = Package.Workbook.Worksheets[1];

            // read document data
            int rowCount = Worksheet.Dimension.End.Row;
            int columnCount = Worksheet.Dimension.End.Column;
            var Cells = Worksheet.Cells;

            // read native and other culture column heading
            data.Cultures = new List<string>();
            for (int col = 3; col <= columnCount; col++)
            {
                if (Cells[1, col] != null)
                {
                    // third column is NativeCulture
                    if (col == 3)
                        data.NativeCulture = Cells[1, col].Text;
                    data.Cultures.Add(Cells[1, col].Text);
                }
            }

            int index = 2;
            int cultureCount = data.Cultures.Count;
            List<InternalRecord> records = new List<InternalRecord>(rowCount / 2);


            // read all translation keys down to <<DO NOT TRANSLATE>>
            for (; Cells[index, 1].Text != serviceData; index++)
            {
                // Create new record and add Key Value
                InternalRecord record = new InternalRecord();
                record.Key = GetKey(Cells[index, 2].Text);

                // Create new Translation list
                record.Translations = new List<InternalText>(cultureCount);

                // Iterate through translations
                for (int culture = 0; culture < cultureCount; culture++)
                {
                    // Create new translation text add current iterated culture and translation
                    InternalText translation = new InternalText();
                    translation.Culture = data.Cultures[culture];
                    translation.Text = SafeMultilineText(Cells[index, culture + 3].Text);

                    // Add cultures translation text to current record
                    record.Translations.Add(translation);
                }
                records.Add(record);
            }

            int indexOfServiceData = index;
            data.Namespaces = new List<InternalNamespace>();
            InternalNamespace lastNS = null;

            index++;
            for (; index < rowCount + 1; index++)
            {
                string source = Cells[index, 1].Text;
                string ns = Cells[index, 2].Text;
                string key = Cells[index, 3].Text;
                string path = Cells[index, 4].Text;

                if (lastNS == null || lastNS.Name != ns)
                {
                    lastNS = new InternalNamespace();
                    lastNS.Name = ns;
                    lastNS.Children = new List<InternalRecord>();
                    data.Namespaces.Add(lastNS);
                }

                InternalRecord record = records[index - indexOfServiceData - 1];
                if (record.Key != key)
                {
                    throw new FormatException("Unexpected key: " + key + "!");
                }

                record.Source = SafeMultilineText(source);
                record.Path = path;
                lastNS.Children.Add(record);
            }

            return data;
        }

        public static void Export_NewDocument(InternalFormat data, string ExcelName)
        {
            // Excel Layout
            Dictionary<string, List<string>> LAYOUT = new Dictionary<string, List<string>>()
            {
                {IMPORT,  new List<string>() },
                {CONTROL, new List<string>(){ ID, KEY, SOURCE, CONTEXT } },
                {CULTURE, new List<string>(){ ID, KEY, SOURCE, CONTEXT, TRANSLATION, DONE, COMMENT} }
            };

            // Create Excel Document and Sheet Map
            var Package = new ExcelPackage();

            // Generate Column Headers
            foreach (string sheetName in LAYOUT.Keys)
            {
                if (sheetName == CULTURE)
                {
                    CreateMultipleWorksheets(ref Package, data.Cultures, LAYOUT[sheetName]);
                }
                else
                {
                    CreateSingleWorksheet(ref Package, sheetName, LAYOUT[sheetName]);
                }
            }

            // Add Data to Sheets
            var rowIndex = 2; // 0: null, 1: columnheaders, 2: first blank row
            foreach (var ns in data.Namespaces)
            {
                foreach (var rec in ns.Children)
                {
                    foreach (string sheetName in LAYOUT.Keys)
                    {
                        if (sheetName == CULTURE)
                        {
                            SetMultipleWorksheetRowsByColumnData(ref Package, data.Cultures, rowIndex, LAYOUT[sheetName], rec);
                        }
                        else
                        {
                            SetSingleWorksheetRowByColumnData(ref Package, sheetName, rowIndex, LAYOUT[sheetName], rec);
                        }
                    }
                    rowIndex++;
                }
            }

            // Saveit
            byte[] ExcelData = Package.GetAsByteArray();
            File.WriteAllBytes(ExcelName, ExcelData);

        }

        public static void Export(InternalFormat data, string ExcelName)
        {

            // Source / Namespace / Key / Directory

            var Package = new ExcelPackage();
            var Worksheet = Package.Workbook.Worksheets.Add("Translation");

            // Caption
            Worksheet.Row(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
            Worksheet.Row(1).Style.Fill.BackgroundColor.SetColor(Color.Orange);
            Worksheet.Row(1).Style.Font.Bold = true;
            Worksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // Establish column headings in cells A1, B1 and other.
            Worksheet.Column(1).Width = 10;
            Worksheet.Cells[1, 1].Value = "#";
            Worksheet.Column(2).Width = 40;
            Worksheet.Cells[1, 2].Value = "ID";
            Worksheet.Column(3).Width = 100;
            Worksheet.Cells[1, 3].Value = data.NativeCulture;
            for (int i = 0, j = 4; i < data.Cultures.Count; i++)
            {
                if (data.Cultures[i] == data.NativeCulture)
                    continue;
                Worksheet.Column(j).Width = 100;
                Worksheet.Cells[1, j].Value = data.Cultures[i];
                j++;
            }

            int index = 2;
            foreach (var ns in data.Namespaces)
                foreach (var rec in ns.Children)
                {
                    Worksheet.Cells[index, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    Worksheet.Cells[index, 1].Value = index - 1;
                    Worksheet.Cells[index, 2].Value = MakeName(ns.Name, rec.Key);
                    Worksheet.Cells[index, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Worksheet.Cells[index, 3].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 229, 212));
                    Worksheet.Cells[index, 3].Style.WrapText = true;
                    Worksheet.Cells[index, 3].Value = rec[data.NativeCulture];
                    for (int i = 0, j = 4; i < data.Cultures.Count; i++)
                    {
                        if (data.Cultures[i] == data.NativeCulture)
                            continue;
                        string translation = rec[data.Cultures[i]];
                        Worksheet.Cells[index, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        if (string.IsNullOrWhiteSpace(translation))
                            Worksheet.Cells[index, j].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 199, 206));
                        else
                            Worksheet.Cells[index, j].Style.Fill.BackgroundColor.SetColor(
                                (j % 2 == 0) ? Color.FromArgb(200, 239, 212) : Color.FromArgb(200, 235, 250));
                        Worksheet.Cells[index, j].Style.WrapText = true;
                        Worksheet.Cells[index, j].Value = translation;
                        j++;
                    }
                    index++;
                }

            Worksheet.Cells[index, 1].Style.Font.Color.SetColor(Color.Red);
            Worksheet.Cells[index, 1].Style.Font.Bold = true;
            Worksheet.Cells[index, 1].Value = serviceData;
            index++;

            foreach (var ns in data.Namespaces)
                foreach (var rec in ns.Children)
                {
                    Worksheet.Row(index).Style.Font.Color.SetColor(Color.LightGray);
                    Worksheet.Cells[index, 1].Value = rec.Source;
                    Worksheet.Cells[index, 2].Value = ns.Name;
                    Worksheet.Cells[index, 3].Value = rec.Key;
                    Worksheet.Cells[index, 4].Value = rec.Path;
                    index++;
                }

            byte[] ExcelData = Package.GetAsByteArray();
            File.WriteAllBytes(ExcelName, ExcelData);
        }

    }
}
