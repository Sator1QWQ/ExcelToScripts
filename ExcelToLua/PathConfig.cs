using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToLua
{
    class PathConfig
    {
        public const string ExcelPath = @"E:\Project\Excel";
        public const string OutputPath = @"E:\Project\UnityProject\CrazyGunplay\CrazyGunplay\Assets\LuaScripts\Configs\Config";
        public const string OutputTextPath = @"E:\Project\UnityProject\CrazyGunplay\CrazyGunplay\Assets\LuaScripts\Configs\Text";
        public const string TextRequireDirectory = @"Configs.Text"; //require文本文件时的路径
        public const int UnExportRow = 3;   //前3行是策划看，不导出的
    }
}
