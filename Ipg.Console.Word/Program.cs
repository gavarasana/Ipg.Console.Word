// See https://aka.ms/new-console-template for more information
using Microsoft.Office.Interop.Word;


Application wordApp = null;
Document document = null;
try
{
    Console.WriteLine("Hello, World!");
    var fileToOpen = args[0];
    if (!File.Exists(fileToOpen))
    {
        Console.WriteLine("Please provide a valid word document path");
        return;
    }
    wordApp = new Application();
    document = wordApp.Documents.Open(FileName: fileToOpen, Visible: false, ReadOnly: true);
    /*
    Console.WriteLine(document.Paragraphs.Count);
    var paragraphs = document.Paragraphs;
    var i = 0;
    foreach (Paragraph paragraph in paragraphs)
    {
        Console.WriteLine($"Para {++i}: \n {paragraph.Range.Text}\n");
    }
    */
    var tables = document.Tables;
    int counter = 0;

    Console.WriteLine($"Total tables: {tables.Count}");
    foreach (Table wordTable in tables)
    {

        Console.WriteLine($"{++counter}) Table Id: {wordTable.Title}\t Total Columns: {wordTable.Columns.Count}\t Total Rows: {wordTable.Rows.Count}");

        Microsoft.Office.Interop.Word.Range range = wordTable.Range;
        for (int i = 1; i <= range.Cells.Count; i++)
        {
            Console.WriteLine($"{i}  - {range.Cells[i].Range.Text}");
        }
    }
    Console.WriteLine("------------------------------------------------------------");

    string? facilityName = tables[1].Range.Cells[4].Range.Text;
    string? facilityTaxId = tables[1].Range.Cells[6].Range.Text;
    bool facilityAdd = !string.IsNullOrEmpty(tables[2].Range.Cells[6].Range.Text);
    string? carrierName = tables[3].Range.Cells[4].Range.Text;
    bool allCptsInclusive = !string.IsNullOrEmpty(tables[4].Range.Cells[6].Range.Text);
    List<string> inclusiveCptCodes = new();

    Console.WriteLine($"Facility Name: {facilityName}");
    Console.WriteLine($"FacilityTaxId: {facilityTaxId}");
    Console.WriteLine($"Facility Add?: {facilityAdd}");
    Console.WriteLine($"Carrier Name: {carrierName}");
    Console.WriteLine($"All CPTs: {allCptsInclusive}");

   

    if (!allCptsInclusive)
    {
        Table cptCodesTable = tables[5];

        Microsoft.Office.Interop.Word.Range cptTableRange = cptCodesTable.Range;
        for (int i = 3; i <= cptTableRange.Cells.Count; i = i + 2)
        {
            var cptCode = cptTableRange.Cells[i].Range.Text.Trim();
            if (!string.IsNullOrEmpty(cptCode)) inclusiveCptCodes.Add(cptCode);

        }
        Console.WriteLine("Inclusive CPTs");
        foreach (var cptCode in inclusiveCptCodes)
        {
            Console.Write($"{cptCode}\t");
        }

        List<string> exclusiveCptCodes = new();
        Table exclusiveCptCodesTable = tables[6];

        Microsoft.Office.Interop.Word.Range exclusiveCptCodesTableRange = exclusiveCptCodesTable.Range;
        for (int i = 3; i <= exclusiveCptCodesTableRange.Cells.Count; i = i + 2)
        {
            var cptCode = exclusiveCptCodesTableRange.Cells[i].Range.Text.Trim();
            if (!string.IsNullOrEmpty(cptCode)) exclusiveCptCodes.Add(cptCode);

        }
        Console.WriteLine("Exclusive CPTs");
        foreach (var cptCode in exclusiveCptCodes)
        {
            Console.Write($"{cptCode}\t");
        }
    }
    /*
    foreach (Table wordTable in tables)
    {
        
        Console.WriteLine($"{++counter}) Table Id: {wordTable.Title}\t Total Columns: {wordTable.Columns.Count}\t Total Rows: {wordTable.Rows.Count}");

        Microsoft.Office.Interop.Word.Range range = wordTable.Range;
        for (int i = 1; i <= range.Cells.Count; i++)
        {
            //Console.WriteLine($"{range.Cells[i].RowIndex} : {range.Cells[i].ColumnIndex} - {range.Cells[i].Range.Text}");
            Console.WriteLine($"{i}  - {range.Cells[i].Range.Text}");

            
            if (range.Cells[i].RowIndex == wordTable.Rows.Count)
            {
                //range.Cells[i].Range.Text = range.Cells[i].RowIndex + ":" + range.Cells[i].ColumnIndex;
                Console.WriteLine(range.Cells[i].Range.Text);
            }
            
        }

        
        ////foreach (Row tableRow in wordTable.Rows)
        ////{
           
        ////    foreach (Cell wordCell in tableRow.Cells)
        ////    {
        ////        //Console.Write($"Text: {wordCell.Range.Text}\t");
        ////        Console.Write($"{wordCell.Row.Range.Text}\t");
        ////    }
        ////    Console.WriteLine("-----");
        ////}

        

    }
    */
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
document?.Close();
Console.WriteLine("Closing application");
wordApp?.Quit();

