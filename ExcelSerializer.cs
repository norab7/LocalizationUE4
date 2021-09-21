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
        private const string serviceData = "--== !!! DO NOT TRANSLATE THE TEXT BELOW !!! == SERVICE DATA ==--";

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

        private static void FormatSheetHeader(ExcelWorksheet sheet)
        {
            if (sheet.Name != "Import")
            {
                // Appearance
                sheet.Row(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Row(1).Style.Fill.BackgroundColor.SetColor(Color.Orange);
                sheet.Row(1).Style.Font.Bold = true;
                sheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // Freeze and Filter
                sheet.View.FreezePanes(2, 1);
                sheet.Cells[sheet.Dimension.Address].AutoFilter = true;
            }

            // AutoWidth
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
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

            // read native and other cultures
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

            // read all translation keys
            for (; Cells[index, 1].Text != serviceData; index++)
            {
                InternalRecord record = new InternalRecord();
                record.Key = GetKey(Cells[index, 2].Text);
                record.Translations = new List<InternalText>(cultureCount);
                for (int culture = 0; culture < cultureCount; culture++)
                {
                    InternalText translation = new InternalText();
                    translation.Culture = data.Cultures[culture];
                    translation.Text = SafeMultilineText(Cells[index, culture + 3].Text);
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
            // Create Excel Document and Sheet Map
            var Package = new ExcelPackage();
            var SheetDictionary = new Dictionary<string, ExcelWorksheet>();

            // Import Sheet
            string ImportName = "Import";
            var ImportSheet = Package.Workbook.Worksheets.Add(ImportName);

            SheetDictionary.Add(ImportName, ImportSheet);

            // Control Sheet
            string ControlName = "Control";
            var ControlSheet = Package.Workbook.Worksheets.Add(ControlName);

            ControlSheet.Cells[1, 1].Value = "Key";
            ControlSheet.Cells[1, 2].Value = "English";
            ControlSheet.Cells[1, 3].Value = "Category";
            ControlSheet.Cells[1, 4].Value = "Context";

            FormatSheetHeader(ControlSheet);
            SheetDictionary.Add(ControlName, ControlSheet);

            // Categories Sheet
            string CategoriesName = "Categories";
            var CategoriesSheet = Package.Workbook.Worksheets.Add(CategoriesName);

            CategoriesSheet.Cells[1, 1].Value = "Category";
            CategoriesSheet.Cells[1, 2].Value = "Description";

            CategoriesSheet.Cells[2, 1].Value = "A0";
            CategoriesSheet.Cells[3, 1].Value = "A1";
            CategoriesSheet.Cells[4, 1].Value = "A2";
            CategoriesSheet.Cells[5, 1].Value = "A3";
            CategoriesSheet.Cells[6, 1].Value = "A4";
            CategoriesSheet.Cells[7, 1].Value = "A5";

            FormatSheetHeader(CategoriesSheet);
            SheetDictionary.Add(CategoriesName, CategoriesSheet);

            // Translation Sheets
            for (int i = 0; i < data.Cultures.Count; i++)
            {
                if (data.Cultures[i] == data.NativeCulture)
                {
                    continue;
                }

                var CultureSheet = Package.Workbook.Worksheets.Add(data.Cultures[i]);

                CultureSheet.Cells[1, 1].Value = "Key";
                CultureSheet.Cells[1, 2].Value = "Category";
                CultureSheet.Cells[1, 3].Value = "Context";
                CultureSheet.Cells[1, 4].Value = "English";
                CultureSheet.Cells[1, 5].Value = "Translation";
                CultureSheet.Cells[1, 6].Value = "Done";
                CultureSheet.Cells[1, 7].Value = "Comment";

                FormatSheetHeader(CultureSheet);
                SheetDictionary.Add(data.Cultures[i], CultureSheet);
            }

            // Add Data Sheets
            var index = 2;
            foreach (var ns in data.Namespaces)
            {
                foreach (var rec in ns.Children)
                {
                    // Fill Import & Control
                    SheetDictionary[ImportName].Cells[index - 1, 1].Value = rec.Key;
                    SheetDictionary[ImportName].Cells[index - 1, 2].Value = rec[data.NativeCulture];
                    SheetDictionary[ControlName].Cells[index, 1].Value = rec.Key;
                    SheetDictionary[ControlName].Cells[index, 2].Value = rec[data.NativeCulture];

                    // Fill Translation Tabs
                    for (int i = 0; i < data.Cultures.Count; i++)
                    {
                        if (data.Cultures[i] == data.NativeCulture)
                        {
                            continue;
                        }
                        ExcelWorksheet sheet = SheetDictionary[data.Cultures[i]];

                        sheet.Cells[index, 1].Value = rec.Key;
                        sheet.Cells[index, 4].Value = rec[data.NativeCulture];
                        sheet.Cells[index, 5].Value = rec[data.Cultures[i]];
                        sheet.Cells[index, 6].Value = "";

                        sheet.Column(1).Hidden = true;
                        sheet.Column(4).Width = 75;
                        sheet.Column(4).Style.WrapText = true;
                        sheet.Column(5).Width = 75;
                        sheet.Column(5).Style.WrapText = true;

                        if (sheet.Name != "de")
                        {
                            continue;
                        }
                        // Format the Done column to be red if it's empty
                        var requiredFormat = sheet.ConditionalFormatting.AddExpression(new ExcelAddress("F" + index));
                        requiredFormat.Style.Fill.BackgroundColor.Color = Color.Red;
                        requiredFormat.Formula = "(F" + index + " = \"\")";

                        var doneFormat = sheet.ConditionalFormatting.AddExpression(new ExcelAddress("F" + index));
                        doneFormat.Style.Fill.BackgroundColor.Color = Color.Green;
                        doneFormat.Formula = "(F" + index + " <> \"\")";

                    }

                    index++;
                }
            }

            // Import will be empty before population, format after
            FormatSheetHeader(ImportSheet);

            // Saveit
            byte[] ExcelData = Package.GetAsByteArray();
            File.WriteAllBytes(ExcelName, ExcelData);

        }

        public static void Export(InternalFormat data, string ExcelName)
        {
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
