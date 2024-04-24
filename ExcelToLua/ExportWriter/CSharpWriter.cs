using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToLua
{
    class CSharpWriter : ExportWriter
    {
        public override string ExportPath => PathConfig.CSharpOutputPath;

        public override bool IsReadOnce => true;

        public override string ValueTab => "";

        public override void Export(string sheetName, string db, Dictionary<string, string> textDic)
        {
            string outputFile = PathConfig.CSharpOutputPath + "\\" + "Config_" + sheetName + ".cs";
            if (File.Exists(outputFile))
            {
                File.Delete(outputFile);
            }

            StreamWriter stream = new StreamWriter(outputFile);
            stream.Write(db);
            stream.Dispose();
        }

        public override string ToScript(string sheetName, int row, string db, string name, string value, string type, string tab)
        {
            string getset = " { get; set; }\n\n";
            if (type.Equals("int"))
            {
                db = db + tab + "int " + name + getset;
            }
            else if(type.Equals("float"))
            {
                db = db + tab + "float " + name + getset;
            }
            else if(type.Equals("bool"))
            {
                db = db + tab + "bool " + name + getset;
            }
            else if(type.Equals("string"))
            {
                db = db + tab + "string " + name + getset;
            }
            else if(type.Contains("array")) //array<int>
            {
                string v = type.Replace("array", "");  //<int>
                string arrayType = v.Substring(1, v.Length - 2);    //去掉尖括号,剩下int
                db = db + tab + "List<" + arrayType + "> " + name + getset;
            }
            else if(type.Contains("enum")) //enum<EType>
            {
                string v = type.Replace("enum", "");  //<EType>
                string enumType = v.Substring(1, v.Length - 2);    //去掉尖括号,剩下EType
                db = db + tab + enumType + " " + name + getset;
            }
            else if(type.Contains("object"))
            {
                db = db + tab + "object " + name + getset;
            }
            return db;
        }

        public override string OnReadSheetStart(ExcelWorksheet sheet, string db, string tab)
        {
            db = db + 
                "/***********************代码由工具生成***********************/\n" +
                "using System.Collections.Generic;\n" +
                "using XLua;\n" +
                "\n" +
                "[CSharpCallLua]\n" +
                "public interface Config_" + sheet.Name + " : IConfigBase\n" +
                "{\n";
            return db;
        }

        public override string OnReadSheetEnd(ExcelWorksheet sheet, string db, string tab)
        {
            db = db + "}";
            return db;
        }
    }
}
