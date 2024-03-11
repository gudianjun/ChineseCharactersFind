// See https://aka.ms/new-console-template for more information
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.Diagnostics;
using System.IO;
using System.Text;
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
//string str = "\r\nステップ①\r\n「電話番号」：080-9811-1a发0000\r\n「納入先名(半角カタカナ)」：ｵｵｼﾞﾏ\r\n「納入先名(全角漢字)」：大島\r\n「〒」：1110036\r\n「府県コード」：東京都\r\n「住所(半角カタカナ)」：ﾄｳｷｮｳﾄﾀｲﾄｳｸ\r\n「住所(全角漢字)」：東京都台東区松が谷\r\n登録を押す\r\n\r\nステップ②\r\n戻るボタンを押す";

//string errStr = string.Empty;
//foreach (var s in str)
//{
//    if (!CheckJpnString(s.ToString(), out errStr))
//    {


//    }
//}

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
var rootCommand = new RootCommand();
var pathOption = new Option<string>
    ("--p", "Xlsx ファイルを含む有効なパス");

rootCommand.Add(pathOption);
string exePath = System.AppDomain.CurrentDomain.BaseDirectory;
string outFileName = Path.Combine(exePath, "ChineseCode.xlsx");
string path = string.Empty;
rootCommand.SetHandler((pathOption) =>
{
    path = pathOption;

}, pathOption);

await rootCommand.InvokeAsync(args);

if (string.IsNullOrEmpty(path))
{
    path = exePath;
}

var xlsxFiles = Directory.GetFiles(path, "*.xlsx", SearchOption.AllDirectories);

List<Tuple<string, List<Tuple<string, string, string, string>>>> logs = new List<Tuple<string, List<Tuple<string, string, string, string>>>>();
foreach (var file in xlsxFiles)
{
    using (var package = new ExcelPackage(new FileInfo(file)))
    {
        string fileName = file;
        List<Tuple<string, string, string, string>> errors = new List<Tuple<string, string, string, string>>();
        foreach (var worksheet in package.Workbook.Worksheets)
        {
            if (worksheet.Dimension == null) { continue; }
            for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
            {
                for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    var text = cell.Text;
                    string errStr = string.Empty;
                    if (!CheckJpnString(text, out errStr))
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Chinese Found => PATH：{file}，sheet：{worksheet.Name}，ROW：{row}，COL：{col}，TEXT:{text}");
                        errors.Add(new Tuple<string, string, string, string>(worksheet.Name, $"[{row},{col}]", text, errStr));
                    }
                }
            }
        }
        Console.ForegroundColor = ConsoleColor.White;
        logs.Add(new Tuple<string, List<Tuple<string, string, string, string>>>(fileName.Replace(path, ""), errors));
    }
}


if (File.Exists(outFileName))
{
    File.Delete(outFileName);
}
using (var package = new ExcelPackage())
{
    var sheet = package.Workbook.Worksheets.Add("LOG");
    sheet.Column(1).Width = 80;
    sheet.Column(2).Width = 30;
    sheet.Column(3).Width = 30;
    sheet.Column(4).Width = 80;
    sheet.Column(5).Width = 80;
    sheet.Cells[1, 1, 1, 5].Merge = true;
    sheet.Cells[1, 1].Style.Font.Bold = true;
    sheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
    sheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(191, 191, 255));
    sheet.Cells[2, 1, 2, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
    sheet.Cells[2, 1, 2, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(191, 215, 255));
    sheet.Cells[2, 1, 2, 5].Style.Font.Bold = true;

    sheet.Cells[1, 1].Value = path;
    sheet.Cells[2, 1].Value = "ファイル名";
    sheet.Cells[2, 2].Value = "SheetName";
    sheet.Cells[2, 3].Value = "Position";
    sheet.Cells[2, 4].Value = "Cell Text";
    sheet.Cells[2, 5].Value = "Error Text";
    int row = 3;
    int col = 1;
    foreach (var log in logs)
    {
        sheet.Cells[row, col++].Value = log.Item1;
        if (log.Item2.Count > 0)
        {
            foreach (var err in log.Item2)
            {
                sheet.Cells[row, col++].Value = err.Item1;
                sheet.Cells[row, col++].Value = err.Item2;
                sheet.Cells[row, col++].Value = err.Item3;
                sheet.Cells[row, col++].Value = err.Item4;
                col = 2;
                row++;
            }
            col = 1;
            continue;
        }
        else
        {
            sheet.Cells[row, col++].Value = "";
            sheet.Cells[row, col++].Value = "";
            sheet.Cells[row, col++].Value = "";
            sheet.Cells[row, col++].Value = "";
        }

        col = 1;
        row++;
    }

    var cell = sheet.Cells[1, 1, row - 1, 5];
    cell.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
    cell.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
    cell.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
    cell.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
    package.SaveAs(outFileName);
}


System.Diagnostics.Process.Start(new ProcessStartInfo
{
    FileName = outFileName,
    UseShellExecute = true
});

bool CheckJpnString(string inputStr, out string errorStr)
{
    StringBuilder stringBuilder = new StringBuilder();
    foreach (var s in inputStr)
    {
        string input = s.ToString();
        int asciiLen = input.Where(c => (int)c <= 127).ToList().Count;
        int doubleLen = input.Length - asciiLen;
        int allLen = asciiLen + doubleLen * 2;

        byte[] bytesUtf8 = Encoding.UTF8.GetBytes(input);

        byte[] shiftJisBytes = Encoding.Convert(Encoding.UTF8, Encoding.GetEncoding(932), bytesUtf8);
        if (bytesUtf8.Length > 1 && shiftJisBytes.Length == 1 && shiftJisBytes[0] == 63)
        {
            stringBuilder.AppendLine(input);
        }
        else
        {
            bool bl = IsJapaneseByte(shiftJisBytes);
            if (!bl)
            {
                stringBuilder.Append(input + ",");
            }
        }

    }
    errorStr = stringBuilder.ToString();
    return errorStr.Length == 0 ? true : false;
}

bool IsJapaneseByte(byte[] bytes)
{
    for (int i = 0; i < bytes.Length; i++)
    {
        byte bt = bytes[i];
        if ((bt >= 0x00 & bt <= 0x1f) || bt == 0x7f)
        {
            continue;
        }
        else if ((bt >= 0x20 & bt <= 0x7e) || (bt >= 0xa1 & bt <= 0xdf))
        {
            continue;
        }
        else if ((bt >= 0x81 & bt <= 0x9f) || (bt >= 0xe0 & bt <= 0xef) && i + 1 < bytes.Length)
        {
            byte btNex = bytes[i + 1];
            if ((btNex >= 0x40 & btNex <= 0x7e) || (btNex >= 0x80 & btNex <= 0xfc))
            {
                i++;
                continue;
            }
            else
            {
                return false;
            }
        }
        else
        {
            return false;
        }
    }

    return true;
}
