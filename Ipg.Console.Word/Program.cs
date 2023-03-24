// See https://aka.ms/new-console-template for more information
using Microsoft.Office.Interop.Word;

const string CAN_BILL_KEYWORD = "which IPG can bill Health Plan for Covered Services";
const string CSV_EXTENSION = ".csv";

Application wordApp = null;
Document document = null;
var outputFile = Path.Combine(Path.GetTempPath(), string.Concat(Path.GetFileNameWithoutExtension(Path.GetRandomFileName()), CSV_EXTENSION));
try
{
    char[] nonPrintableChars = new char[] { '\r', '\n', '\a' };

    var folderPath = args[0];
    if (!Directory.Exists(folderPath))
    {
        Console.WriteLine("Please provide a valid folder path");
        return;
    }

    using (StreamWriter csvFile = new StreamWriter(outputFile, true))
    {
        csvFile.WriteLine("Facility Name|Facility TaxId|Carrier Name|CPT Codes|CPT Desc|CPT Included or Excluded|Effective Date");
    }

    wordApp = new Application();

    var files = Directory.GetFiles(folderPath, "*.docx");
    foreach (var file in files)
    {
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"Processing file {file}");
        document = wordApp.Documents.Open(FileName: file, Visible: false, ReadOnly: true);


        var tables = document.Tables;
        int totalTables = tables.Count;

        string? facilityName = tables[1].Range.Cells[4].Range.Text.Trim(nonPrintableChars);
        string? facilityTaxId = tables[1].Range.Cells[6].Range.Text.Trim(nonPrintableChars);

        string? carrierName = tables[3].Range.Cells[4].Range.Text.Trim(nonPrintableChars);
        bool allCptsInclusive = !string.IsNullOrEmpty(tables[4].Range.Cells[6].Range.Text.Trim(nonPrintableChars));
        string? effectiveDate = tables[totalTables].Range.Cells[4].Range.Text.Trim(nonPrintableChars);

        // No need to get CPT codes, since all CPT codes are included
        if (allCptsInclusive)
        {
            using (StreamWriter csvFile = new StreamWriter(outputFile, true))
            {
                csvFile.WriteLine($"{facilityName}|{facilityTaxId}|{carrierName}|All||Included|{effectiveDate}");
            }
        }
        else
        {
            Dictionary<string, string> cptCodes = new();
            bool inclusiveCptsSpecified = false;

            Paragraphs paragraphs = document.Paragraphs;
            foreach (Paragraph paragraph in paragraphs)
            {
                if (paragraph.Range.Text.Contains(CAN_BILL_KEYWORD, StringComparison.CurrentCultureIgnoreCase))
                {
                    inclusiveCptsSpecified = true;
                    break;
                }
            }

            if (inclusiveCptsSpecified)
            {
                cptCodes = GetCptCodes(nonPrintableChars, tables[5].Range);
                using StreamWriter csvFile = new(outputFile, true);
                foreach (var cptCode in cptCodes)
                {
                    csvFile.WriteLine($"{facilityName}|{facilityTaxId}|{carrierName}|{cptCode.Key}|{cptCode.Value}|Included|{effectiveDate}");
                }
            }

            // Inclusive CPT codes not found. Lets get exclusive CPT codes.
            // In some word documents, there is only one table, which could either be inclusive or exclusive.
            if (cptCodes.Count == 0)
            {
                Microsoft.Office.Interop.Word.Range excludeCptTableRange = (totalTables > 8) ? tables[6].Range : tables[5].Range;
                cptCodes = GetCptCodes(nonPrintableChars, excludeCptTableRange);
                using StreamWriter csvFile = new(outputFile, true);
                foreach (var cptCode in cptCodes)
                {
                    csvFile.WriteLine($"{facilityName}|{facilityTaxId}|{carrierName}|{cptCode.Key}|{cptCode.Value}|Excluded|{effectiveDate}");
                }
            }
        }

        // Close word document
        document.Close();
        document = null;
    }
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
finally
{
    document?.Close();
    document = null;
    wordApp?.Quit();
    wordApp = null;
    var originalColor = Console.ForegroundColor;
    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine($"Please check outputfile {outputFile} for results");
    Console.ForegroundColor = originalColor;
    Console.WriteLine("Press any key to exit");
    Console.ReadKey();
}

static Dictionary<string, string> GetCptCodes(char[] nonPrintableChars, Microsoft.Office.Interop.Word.Range cptTableRange)
{
    Dictionary<string, string> cptCodes = new();

    for (int i = 3; i <= cptTableRange.Cells.Count; i += 2)
    {
        var cptCode = cptTableRange.Cells[i].Range.Text.Trim(nonPrintableChars);
        var cptDescription = cptTableRange.Cells[i + 1].Range.Text.Trim(nonPrintableChars);
        if (!string.IsNullOrEmpty(cptCode) && !cptCodes.ContainsKey(cptCode)) cptCodes.Add(cptCode, cptDescription);

    }
    return cptCodes;
}

// TO BE DELETED LATER

//Console.WriteLine($"Facility Name: {facilityName}");
//Console.WriteLine($"FacilityTaxId: {facilityTaxId}");
//Console.WriteLine($"Facility Add?: {facilityAdd}");
//Console.WriteLine($"Carrier Name: {carrierName}");
//Console.WriteLine($"All CPTs: {allCptsInclusive}");
//Console.WriteLine($"Effective Date: {effectiveDate}");
//foreach (Table wordTable in tables)
//{

//    Console.WriteLine($"{++counter}) Table Id: {wordTable.Title}\t Total Columns: {wordTable.Columns.Count}\t Total Rows: {wordTable.Rows.Count}");

//    Microsoft.Office.Interop.Word.Range range = wordTable.Range;
//    for (int i = 1; i <= range.Cells.Count; i++)
//    {
//        Console.WriteLine($"{i}  - {range.Cells[i].Range.Text}");
//    }
//}
//Console.WriteLine("------------------------------------------------------------");

////Console.WriteLine(document.Paragraphs.Count);
//      //var paragraphs = document.Paragraphs;
//      //var i = 0;
//      //foreach (Paragraph paragraph in paragraphs)
//      //{
//      //    Console.WriteLine($"Para {++i}: \n {paragraph.Range.Text}\n");
//      //}



//      foreach (Table wordTable in tables)
//      {

//          Console.WriteLine($"{++counter}) Table Id: {wordTable.Title}\t Total Columns: {wordTable.Columns.Count}\t Total Rows: {wordTable.Rows.Count}");

//          Microsoft.Office.Interop.Word.Range range = wordTable.Range;
//          for (int i = 1; i <= range.Cells.Count; i++)
//          {
//              //Console.WriteLine($"{range.Cells[i].RowIndex} : {range.Cells[i].ColumnIndex} - {range.Cells[i].Range.Text}");
//              Console.WriteLine($"{i}  - {range.Cells[i].Range.Text}");


//              if (range.Cells[i].RowIndex == wordTable.Rows.Count)
//              {
//                  //range.Cells[i].Range.Text = range.Cells[i].RowIndex + ":" + range.Cells[i].ColumnIndex;
//                  Console.WriteLine(range.Cells[i].Range.Text);
//              }

//          }


//          ////foreach (Row tableRow in wordTable.Rows)
//          ////{

//          ////    foreach (Cell wordCell in tableRow.Cells)
//          ////    {
//          ////        //Console.Write($"Text: {wordCell.Range.Text}\t");
//          ////        Console.Write($"{wordCell.Row.Range.Text}\t");
//          ////    }
//          ////    Console.WriteLine("-----");
//          ////}



//      }

//bool facilityAdd = !string.IsNullOrEmpty(tables[2].Range.Cells[6].Range.Text.Trim(nonPrintableChars));