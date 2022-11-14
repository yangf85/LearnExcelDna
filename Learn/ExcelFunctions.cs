using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Learn
{
    public static class ExcelFunctions
    {
        [ExcelFunction(Name = nameof(ArrayFunction), Category = "Learn")]
        public static object[,] ArrayFunction()
        {
            var array = new object[5, 5];
            var random = new Random();
            for (int i = 0; i < 5; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    array[i, j] = random.Next(5, 100);
                }
            }

            return array;
        }

        [ExcelFunction(Name = nameof(CalculateArea), Description = "年轻人的第一个Excel插件", Category = "Learn")]
        public static object CalculateArea([ExcelArgument(Name = "参数名称", Description = "面积参数名称")] object[] list)
        {
            var flag = list[0]?.ToString().ToUpper();
            switch (flag)
            {
                case "R":
                    return (double)list[1] * (double)list[2];

                case "C":
                    return Math.PI * (double)list[1] * (double)list[1];

                case "T":
                    return ((double)list[1] + (double)list[2]) * (double)list[3] / 2;

                default:
                    return ExcelError.ExcelErrorValue;
            }
        }

        [ExcelCommand(Name = "Button", MenuName = "按钮", MenuText = "点我")]
        public static void CommandTest()
        {
            var s = "";
        }

        [ExcelFunction(Name = nameof(SayHello), Description = "年轻人的第一个Excel插件", Category = "Learn")]
        public static string SayHello()
        {
            return "Hello Excel";
        }
    }
}