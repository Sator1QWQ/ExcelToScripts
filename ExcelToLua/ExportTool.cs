using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToLua
{
    class ExportTool
    {
        /// <summary>
        /// 将数组根据逗号拆分并且去掉一层中括号。
        /// 比如arr为[2,[5,6],8]这种格式的，调用后结果为string[3] = 2 [5,6] 8
        /// </summary>
        /// <param name="arr"></param>
        /// <returns></returns>
        public static string[] SplitArray(string arr)
        {
            int startIndex = 1;
            int count = (arr.Length - 1) - startIndex;
            string arr2 = arr.Substring(startIndex, count); //2,[[5],[6]],8
            List<string> split = new List<string>();

            int index = 0;
            while (index < arr2.Length)
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
                    if (rightIndex < arr2.Length - 1)
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
                    if (index == arr2.Length)
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
            for (int i = leftIndex; i < str.Length; i++)
            {
                char c = str[i];
                if (c == '[')
                {
                    stack.Push(i);
                }
                else if (c == ']')
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
