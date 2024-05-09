using ClosedXML.Excel;
using ConsoleApp1;

Console.WriteLine("file name: ");
string fromFile = Console.ReadLine();

Console.WriteLine("code extension (with '.'): ");
string codeExtension = Console.ReadLine();

Console.WriteLine("reading file...");
XLWorkbook workbook = new XLWorkbook(fromFile);
Console.WriteLine("ok");

string sheetName = "";
while (sheetName != "0")
{
    Console.WriteLine("sheet name (input '0' to exit): ");
    sheetName = Console.ReadLine();
    if (sheetName != "0")
    {
        IXLWorksheet inputSheet = workbook.Worksheet(sheetName);
        Analize(workbook, inputSheet);
    }
}

Console.WriteLine("saving...");

workbook.SaveAs(Path.Combine(Path.GetDirectoryName(fromFile), Path.GetFileNameWithoutExtension(fromFile) + "_new.xlsx"));
workbook.Dispose();

Console.WriteLine("ok");

void Analize(XLWorkbook workbook, IXLWorksheet sheet)
{
    List<FileBlock> fileBlocks = new List<FileBlock>();

    int lastRowNum = sheet.LastRowUsed().RowNumber();
    int lastColNum = sheet.LastColumnUsed().ColumnNumber();
    bool inGroup = false;
    FileBlock fileBlock = null;
    Block block = null;

    Console.WriteLine("analyzing...");

    for (int i = 1; i <= lastRowNum; ++i)
    {
        string value = "";
        if (!sheet.Column(1).Cell(i).IsEmpty())
        {
            value = sheet.Column(1).Cell(i).GetText();
        }
        if (value.EndsWith(codeExtension + ":"))
        {
            if (fileBlock != null)
            {
                fileBlocks.Add(fileBlock);
            }
            fileBlock = new FileBlock();
            fileBlock.blockList = new List<Block>();
            fileBlock.fileName = value.Replace(":", "");
            continue;
        }

        if (value != "")
        {
            if (!inGroup)
            {
                block = new Block();
                block.rows = new List<string>();
                inGroup = true;
            }

            string row = "";
            for (int j = 1; j <= lastColNum; ++j)
            {
                if (!sheet.Column(j).Cell(i).IsEmpty())
                {
                    row += sheet.Column(j).Cell(i).GetText();
                }
            }
            block.rows.Add(row);
        }
        else
        {
            if (inGroup)
            {
                inGroup = false;

                fileBlock.blockList.Add(block);
            }
        }
    }

    Console.WriteLine("writing...");

    string name = sheet.Name;
    workbook.Worksheet(name).Delete();
    IXLWorksheet newSheet = workbook.AddWorksheet(name);
    List<string> options = new List<string>() { "〇", "×" };
    string optionsStr = $"\"{String.Join(",", options)}\"";

    int fileStartRow = 1;
    int nowRow = 1;

    int fileNameCol = 1;
    int indexCol = 2;
    int enableCol = 3;
    int doneCol = 4;
    int blockStrCol = 5;

    int blockIndex = 0;

    for (int i = 0; i < fileBlocks.Count; ++i)
    {
        if (nowRow > 1)
        {
            newSheet.Range(fileStartRow, 1, nowRow - 1, 1).Merge();
        }

        fileStartRow = nowRow;

        fileBlock = fileBlocks[i];

        newSheet.Cell(nowRow, fileNameCol).SetValue(fileBlock.fileName);

        for (int j = 0; j < fileBlock.blockList.Count; ++j)
        {
            block = fileBlock.blockList[j];
            newSheet.Cell(nowRow, indexCol).SetValue((++blockIndex).ToString());
            newSheet.Cell(nowRow, enableCol).CreateDataValidation();
            newSheet.Cell(nowRow, enableCol).GetDataValidation().AllowedValues = XLAllowedValues.List;
            newSheet.Cell(nowRow, enableCol).GetDataValidation().InCellDropdown = true;
            newSheet.Cell(nowRow, enableCol).GetDataValidation().List(optionsStr, true);
            newSheet.Cell(nowRow, doneCol).CreateDataValidation();
            newSheet.Cell(nowRow, doneCol).GetDataValidation().AllowedValues = XLAllowedValues.List;
            newSheet.Cell(nowRow, doneCol).GetDataValidation().InCellDropdown = true;
            newSheet.Cell(nowRow, doneCol).GetDataValidation().List(optionsStr, true);

            string blockStr = string.Join(" \n", block.rows);
            newSheet.Cell(nowRow, blockStrCol).SetValue(blockStr);

            ++nowRow;
        }
    }
    newSheet.Range(fileStartRow, 1, nowRow - 1, 1).Merge();

    Console.WriteLine("setting color...");

    lastRowNum = newSheet.LastRowUsed().RowNumber();
    List<string> searchList = new List<string>() { ".ColumnCount - 1", ".RowCount - 1", "Cdate", "Format", "yymmdd" };
    for (int i = 1; i <= lastRowNum; ++i)
    {
        foreach (string str in searchList)
        {
            int index = 0;
            if (!newSheet.Cell(i, 5).IsEmpty() && newSheet.Cell(i, 5).GetText().ToLower().Contains(str.ToLower()))
            {
                while (index > -1)
                {
                    index = newSheet.Cell(i, 5).GetText().ToLower().IndexOf(str.ToLower(), index);
                    if (index > -1)
                    {
                        newSheet.Cell(i, 5).GetRichText().Substring(index, str.Length).SetBold().SetFontColor(XLColor.Red).SetUnderline().SetShadow(true);
                        index += str.Length;
                    }
                }
            }
        }
    }

    Console.WriteLine("setting style...");

    newSheet.Row(1).InsertRowsAbove(1);
    newSheet.Cell(1, 1).SetValue("ファイル");
    newSheet.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.LightGreen;
    newSheet.Cell(1, 2).SetValue("インデックス");
    newSheet.Cell(1, 2).Style.Fill.BackgroundColor = XLColor.LightGreen;
    newSheet.Cell(1, 3).SetValue("改修要");
    newSheet.Cell(1, 3).Style.Fill.BackgroundColor = XLColor.LightGreen;
    newSheet.Cell(1, 4).SetValue("完了");
    newSheet.Cell(1, 4).Style.Fill.BackgroundColor = XLColor.LightGreen;
    newSheet.Cell(1, 5).SetValue("コード");
    newSheet.Cell(1, 5).Style.Fill.BackgroundColor = XLColor.LightGreen;

    newSheet.Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
    newSheet.Column(1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
    newSheet.Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    newSheet.Column(2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    newSheet.Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    newSheet.Column(3).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    newSheet.Column(4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    newSheet.Column(4).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

    newSheet.Column(1).AdjustToContents();
    newSheet.Column(2).Width = 12;
    newSheet.Column(3).Width = 7;
    newSheet.Column(4).Width = 7;
    newSheet.Column(5).AdjustToContents();
    newSheet.Column(5).Style.Alignment.WrapText = true;

    newSheet.Range(1, 1, nowRow, 5).SetAutoFilter();

    newSheet.Range(1, 1, nowRow, 5).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
    newSheet.Range(1, 1, nowRow, 5).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
}