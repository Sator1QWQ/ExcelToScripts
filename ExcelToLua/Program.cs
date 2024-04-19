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
                                dbTable = ToLua(sheet.Name, row, dbTable, cellName, cellValue, cellType, tab + tab);

                                if (cellType.Equals("string"))
                                {
                                    textDic.Add(cellName + "_" + row.ToString(), cellValue);
                                }
                            }
                            dbTable = exportWriter.OnReadRowEnd(dbTable, tab);
                        }

                        dbTable = exportWriter.OnReadSheetEnd(sheet, dbTable, tab);
                        exportWriter.Export(sheet.Name, dbTable, textDic);
                        Console.WriteLine("[" + sheet.Name + "]导出成功");
                    }
                }
            }
        }

        //把一个数据转换成lua
        private static string ToLua(string sheetName, int row, string db, string name, string value, string type, string tab)
        {
            if(type.Equals("number"))
            {
                db = db + tab + name + " = " + value + ",\n";
            }
            else if(type.Equals("bool"))
            {
                string newValue = "";
                if(value.Equals("1"))
                {
                    newValue = "true";
                }
                else if(value.Equals("0"))
                {
                    newValue = "false";
                }
                db = db + tab + name + " = " + newValue + ",\n";
            }
            else if(type.Equals("string"))
            {
                db = db + tab + name + " = " + sheetName + "_Text." + name + "_" + row.ToString() + ",\n";
            }
            else if(type.Equals("array"))
            {
                string[] arr = null;
                if (string.IsNullOrEmpty(value))
                {
                    arr = new string[] { };
                }
                else
                {
                    arr = ExportTool.SplitArray(value);
                }
                string nextTab = tab + "	";
                db = db + tab + name + " = {\n";
                for(int i = 0; i < arr.Length; i++)
                {
                    if (arr[i].StartsWith("["))
                    {
                        db = ToLua(sheetName, row, db, "[" + (i + 1).ToString() + "]", arr[i], "array", nextTab);
                    }
                    else
                    {
                        db = db + nextTab + "[" + (i+1) + "] = " + arr[i] + ",\n";
                    }
                }
                db = db + tab + "},\n";
            }
            return db;
        }

        
    }
}
