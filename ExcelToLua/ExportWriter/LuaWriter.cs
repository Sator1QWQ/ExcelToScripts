using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToLua
{
    /// <summary>
    /// 导出lua表
    /// </summary>
    class LuaWriter : ExportWriter
    {
        public override string ExportPath => PathConfig.LuaOutputPath;
        public string TextExportPath => PathConfig.LuaTextOutputPath;

        public override string OnReadRowStart(string id, string db, string tab)
        {
            db = db + tab + "[" + id + "] = {\n";
            return db;
        }

        public override string OnReadRowEnd(string db, string tab)
        {
            db = db + tab + "},\n";
            return db;
        }

        public override string OnReadSheetEnd(ExcelWorksheet sheet, string db, string tab)
        {
            db = db + "}\n";
            db = db + sheet.Name + ".Count = " + (sheet.Dimension.Rows - PathConfig.UnExportRow);
            return db;
        }

        public override void Export(string sheetName, string db, Dictionary<string, string> textDic)
        {
            string textName = sheetName + "_Text";
            string outputFile = PathConfig.LuaOutputPath + "\\" + sheetName + ".lua";
            if (File.Exists(outputFile))
            {
                File.Delete(outputFile);
            }

            StreamWriter stream = new StreamWriter(outputFile);
            stream.WriteLine($"require \"{PathConfig.TextRequireDirectory}.{textName}\"");
            stream.Write(db);
            stream.Dispose();

            //文本单独创建一个文件
            string outputTextFile = PathConfig.LuaTextOutputPath + "\\" + textName + ".lua";
            if (File.Exists(outputTextFile))
            {
                File.Delete(outputTextFile);
            }
            StreamWriter textStream = new StreamWriter(outputTextFile);
            textStream.WriteLine(textName + " = {}");
            foreach (KeyValuePair<string, string> pair in textDic)
            {
                textStream.WriteLine($"{textName}.{pair.Key} = \"{pair.Value}\"");
            }
            textStream.Dispose();
        }

        public override string ToScript(string sheetName, int row, string db, string name, string value, string type, string tab)
        {
            if (type.Equals("number"))
            {
                db = db + tab + name + " = " + value + ",\n";
            }
            else if (type.Equals("bool"))
            {
                string newValue = "";
                if (value.Equals("1"))
                {
                    newValue = "true";
                }
                else if (value.Equals("0"))
                {
                    newValue = "false";
                }
                db = db + tab + name + " = " + newValue + ",\n";
            }
            else if (type.Equals("string"))
            {
                db = db + tab + name + " = " + sheetName + "_Text." + name + "_" + row.ToString() + ",\n";
            }
            else if (type.Equals("array"))
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
                for (int i = 0; i < arr.Length; i++)
                {
                    if (arr[i].StartsWith("["))
                    {
                        db = ToScript(sheetName, row, db, "[" + (i + 1).ToString() + "]", arr[i], "array", nextTab);
                    }
                    else
                    {
                        db = db + nextTab + "[" + (i + 1) + "] = " + arr[i] + ",\n";
                    }
                }
                db = db + tab + "},\n";
            }
            return db;
        }
    }
}
