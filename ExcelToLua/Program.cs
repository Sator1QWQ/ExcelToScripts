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
        static void Main(string[] args)
        {
            ReadExcel();
            Console.Read();
        }

        private static void ReadExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string[] files = Directory.GetFiles(PathConfig.ExcelPath, "*.xlsx", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
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
                    string dbTable = sheet.Name + " = {\n";
                    string tab = "	";

                    //从第4行开始，前面3行是注释行
                    //这里的索引是从1开始，与Excel可以保持一致
                    for (int row = PathConfig.UnExportRow + 1; row <= sheet.Dimension.Rows; row++)
                    {
                        string id = sheet.Cells[row, 1].Value.ToString();
                        dbTable = dbTable + tab + "[" + id + "] = {\n";

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

                            if(cellType.Equals("string"))
                            {
                                textDic.Add(cellName + "_" + row.ToString(), cellValue);
                            }
                        }

                        dbTable = dbTable + tab + "},\n";
                    }

                    dbTable = dbTable + "}\n";
                    dbTable = dbTable + sheet.Name + ".Count = " + (sheet.Dimension.Rows - PathConfig.UnExportRow);

                    string textName = sheet.Name + "_Text";
                    string outputFile = PathConfig.OutputPath + "\\" + sheet.Name + ".lua";
                    if (File.Exists(outputFile))
                    {
                        File.Delete(outputFile);
                    }
                    StreamWriter stream = new StreamWriter(outputFile);
                    stream.WriteLine($"require \"{PathConfig.TextRequireDirectory}.{textName}\"");
                    stream.Write(dbTable);
                    stream.Dispose();

                    //文本单独创建一个文件
                    string outputTextFile = PathConfig.OutputTextPath + "\\" + textName + ".lua";
                    if(File.Exists(outputTextFile))
                    {
                        File.Delete(outputTextFile);
                    }
                    StreamWriter textStream = new StreamWriter(outputTextFile);
                    textStream.WriteLine(textName + " = {}");
                    foreach(KeyValuePair<string, string> pair in textDic)
                    {
                        textStream.WriteLine($"{textName}.{pair.Key} = \"{pair.Value}\"");
                    }
                    textStream.Dispose();

                    Console.WriteLine("[" + sheet.Name + "]导出成功");
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
                    arr = SplitArray(value);
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

        //arr为[2,[5,6],8]这种格式的
        private static string[] SplitArray(string arr)
        {
            int startIndex = 1;
            int count = (arr.Length - 1) - startIndex;
            string arr2 = arr.Substring(startIndex, count); //2,[[5],[6]],8     //111,231,532,666
            List<string> split = new List<string>();

            int index = 0;
            while(index < arr2.Length)
            {
                int removeCount = 0;

                //都是从0截取的，index代表的也是数量
                if (arr2[index] == ',')
                {
                    split.Add(arr2.Substring(0, index));
                    removeCount = index + 1;
                }
                else if (arr2[index] == '[') //遇到[的时候，index只会是0
                {
                    int rightIndex = GetRightIndex(arr2, index);
                    split.Add(arr2.Substring(0, rightIndex + 1));
                    removeCount = rightIndex + 1;
                    
                    //]不是最后一个位置，则需要多移除一个，为了移除]后的逗号
                    if(rightIndex < arr2.Length - 1)
                    {
                        removeCount++;
                    }
                }

                if (removeCount > 0)
                {
                    arr2 = arr2.Remove(0, removeCount);
                    index = 0;
                }
                else
                {
                    index++;

                    //最后一个数
                    if(index == arr2.Length)
                    {
                        split.Add(arr2.Substring(0, arr2.Length));
                        break;
                    }
                }
            }

            return split.ToArray();
        }

        /// <summary>
        /// 获取右括号的索引
        /// </summary>
        /// <param name="str">字符串</param>
        /// <param name="leftIndex">左括号索引</param>
        /// <returns></returns>
        private static int GetRightIndex(string str, int leftIndex)
        {
            //[[[2],5],[6]]

            Stack<int> stack = new Stack<int>();
            for(int i = leftIndex; i < str.Length; i++)
            {
                char c = str[i];
                if (c == '[')
                {
                    stack.Push(i);
                }
                else if(c == ']')
                {
                    stack.Pop();
                    //栈空时，匹配完成
                    if (stack.Count == 0)
                    {
                        return i;
                    }
                }
            }

            //括号不匹配
            return -1;
        }
    }
}
