using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToLua
{
    abstract class ExportWriter
    {
        /// <summary>
        /// 导出路径
        /// </summary>
        public abstract string ExportPath { get; }

        /// <summary>
        /// 每行excel数据转换成代码
        /// </summary>
        /// <param name="sheetName">excel sheet</param>
        /// <param name="row">当前第几行</param>
        /// <param name="db">字符串 最后需要将这个字符串返回</param>
        /// <param name="name">字段名称</param>
        /// <param name="value">字段值</param>
        /// <param name="type">字段类型</param>
        /// <param name="tab">\t</param>
        /// <returns></returns>
        public abstract string ToScript(string sheetName, int row, string db, string name, string value, string type, string tab);

        /// <summary>
        /// 导出操作
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="db"></param>
        /// <param name="textDic">文本字典,key:文本名称,value:文本内容</param>
        public abstract void Export(string sheetName, string db, Dictionary<string, string> textDic);

        /// <summary>
        /// 每一行刚读取时
        /// </summary>
        /// <param name="db"></param>
        /// <returns>转换后的字符串</returns>
        public virtual string OnReadRowStart(string id, string db, string tab) => db;

        /// <summary>
        /// 每一行读取结束时
        /// </summary>
        /// <param name="db"></param>
        /// <returns>转换后的字符串</returns>
        public virtual string OnReadRowEnd(string db, string tab) => db;

        /// <summary>
        /// 每个sheet刚读取时
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="db"></param>
        /// <returns>转换后的字符串</returns>
        public virtual string OnReadSheetStart(ExcelWorksheet sheet, string db) => db;

        /// <summary>
        /// 每个sheet读取结束时
        /// </summary>
        /// <param name="db"></param>
        /// <returns>转换后的字符串</returns>
        public virtual string OnReadSheetEnd(ExcelWorksheet sheet, string db, string tab) => db;
    }
}
