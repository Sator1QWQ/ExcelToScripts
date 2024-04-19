using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExcelToLua
{
    class Program
    {
        private static List<ExportWriter> writerList = new List<ExportWriter>();

        static void Main(string[] args)
        {
            InitWriter();
            ReadExcel();
            Console.Read();
        }

        private static void InitWriter()
        {
            writerList.Add(new LuaWriter());
            writerList.Add(new CSharpWriter());
        }

        private static void ReadExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string[] files = Directory.GetFiles(PathConfig.ExcelPath, "*.xlsx", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                for(int writerIndex = 0; writerIndex < writerList.Count; writerIndex++)
                {
                    ExportWriter exportWriter = writerList[writerIndex];
                    string file = files[i];

                    //正在运行的excel
                    if (file.IndexOf("~$") != -1)
                    {
                        continue;
                    }

                    ExcelPackage package = new ExcelPackage(file);
                    ExcelWorksheets sheets = package.Workbook.Worksheets;
                    foreach (ExcelWorksheet sheet in sheets)
                    {
                        if (sheet.Dimension == null)
                        {
                            Console.WriteLine("存在为空的表！file==" + file + ", sheet==" + sheet.Name);
                            continue;
                        }

                        Dictionary<string, string> textDic = new Dictionary<string, string>();
                        int count = sheet.Dimension.Rows;
                        string dbTable = "";
                        string tab = "	";
                        dbTable = exportWriter.OnReadSheetStart(sheet, dbTable, tab);

                        //从第4行开始，前面3行是注释行
                        //这里的索引是从1开始，与Excel可以保持一致
                        for (int row = PathConfig.UnExportRow + 1; row <= sheet.Dimension.Rows; row++)
                        {
                            //只读取一行数据，下一行直接break掉
                            if(exportWriter.IsReadOnce && row > PathConfig.UnExportRow + 1)
                            {
                                break;
                            }
                            string id = sheet.Cells[row, 1].Value.ToString();
                            dbTable = exportWriter.OnReadRowStart(id, dbTable, tab);

                            for (int col = 1; col <= sheet.Dimension.Columns; col++)
                            {
                                string nowCell = ((char)(col + 64)).ToString() + row;

                                //存在#则不转换这一列
                                if (sheet.Cells[1, col].Value == null || sheet.Cells[1, col].Value.ToString().IndexOf("#") != -1)
                                {
                                    continue;
                                }

                                string cellName = sheet.Cells[2, col].Value.ToString();
                                string cellType = sheet.Cells[3, col].Value.ToString();
                                ExcelRange range = sheet.Cells[row, col];
                                string cellValue = range.Value == null ? "" : range.Value.ToString();
                                dbTable = exportWriter.ToScript(sheet.Name, row, dbTable, cellName, cellValue, cellType, tab + exportWriter.ValueTab);

                                if (cellType.Equals("string"))
                                {
                                    textDic.Add(cellName + "_" + row.ToString(), cellValue);
                                }
                            }
                            dbTable = exportWriter.OnReadRowEnd(dbTable, tab);
                        }

                        dbTable = exportWriter.OnReadSheetEnd(sheet, dbTable, tab);
                        exportWriter.Export(sheet.Name, dbTable, textDic);
                        Console.WriteLine(exportWriter + " [" + sheet.Name + "]导出成功");
                    }
                }
            }
        }
    }
}
