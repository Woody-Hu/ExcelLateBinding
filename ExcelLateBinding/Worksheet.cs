using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLateBinding
{
    /// <summary>
    /// WorkSheet
    /// </summary>
    public class Worksheet
    {
        /// <summary>
        /// workSheet组件
        /// </summary>
        object m_workSheet;

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="thisObject"></param>
        internal Worksheet(object thisObject)
        {
            m_workSheet = thisObject;
        }

        /// <summary>
        /// 工作表的名称
        /// </summary>
        public string Name
        {
            set
            {
                ExcelUtilityMethod.SetProperty(m_workSheet, "Name", new object[] { value });
            }
            get
            {
                return (string)ExcelUtilityMethod.GetProperty(m_workSheet, "Name");
            }
        }

        /// <summary>
        /// 工作表所在的工作簿
        /// </summary>
        public Workbook Parent
        {
            get
            {
                object tempObject = ExcelUtilityMethod.GetProperty(m_workSheet, "Parent");
                return new Workbook(tempObject);
            }
        }

        /// <summary>
        /// 获取一个Range
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public Range Range(string input)
        {
            return new Range(ExcelUtilityMethod.GetProperty(m_workSheet, "Range", new object[] { input, Missing.Value }));
        }

        /// <summary>
        /// 获取一个Range
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public Range Range(int row, int column)
        {
            string rangStr = ExcelUtilityMethod.ConvertToTitle(column + 1) + (row + 1).ToString();
            object tenmpRange = ExcelUtilityMethod.GetProperty(m_workSheet, "Range",
                new object[] { rangStr });
            return new Range(tenmpRange);
        }

        /// <summary>
        /// 获取一个Range
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public Range Range(int rowOne, int columnOne, int rowTwo, int columnTwo)
        {
            string rangStrOne = ExcelUtilityMethod.ConvertToTitle(columnOne + 1) + (rowOne + 1).ToString();
            string rangStrTwo = ExcelUtilityMethod.ConvertToTitle(columnTwo + 1) + (rowTwo + 1).ToString();
            string rangStr = rangStrOne + ":" + rangStrTwo;
            object tenmpRange = ExcelUtilityMethod.GetProperty(m_workSheet, "Range",
                new object[] { rangStr });
            return new Range(tenmpRange);
        }

        /// <summary>
        /// 获取一个Range
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public Range get_Range(string input)
        {
            return Range(input);
        }

        /// <summary>
        /// 获取一个Cell
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public void Cell(int row, int column, dynamic input)
        {
            string rangStr = ExcelUtilityMethod.ConvertToTitle(column + 1) + (row + 1).ToString();
            object tenmpRange = ExcelUtilityMethod.GetProperty(m_workSheet, "Range",
                new object[] { rangStr });
            ExcelUtilityMethod.SetProperty(tenmpRange, "Value", new object[] { input });
        }

        /// <summary>
        /// 获取一个Cell
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        public dynamic Cell(int row, int column)
        {
            string rangStr = ExcelUtilityMethod.ConvertToTitle(column + 1) + (row + 1).ToString();
            object tenmpRange = ExcelUtilityMethod.GetProperty(m_workSheet, "Range",
                new object[] { rangStr });
            return ExcelUtilityMethod.GetProperty(tenmpRange, "Value");
        }

        /// <summary>
        /// 选择此工作表
        /// </summary>
        public void Select()
        {
            ExcelUtilityMethod.UseMethod(m_workSheet, "Select", null);
            return;
        }

        /// <summary>
        /// 在此工作表添加一个图片
        /// </summary>
        /// <param name="strTempPath"></param>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <param name="tempImageHeight"></param>
        /// <param name="tempImageWidth"></param>
        /// <returns></returns>
        public Shape AddPicture(string strTempPath, float left, float top, float tempImageHeight, float tempImageWidth)
        {
            object tempShapes = ExcelUtilityMethod.GetProperty(m_workSheet, "Shapes");
            object temp = ExcelUtilityMethod.UseMethod(tempShapes, "AddPicture", new object[]{strTempPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,
                                left, top, tempImageWidth, tempImageHeight});

            return new Shape(temp);

        }

        /// <summary>
        /// 在此工作表添加一个图表
        /// </summary>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <returns></returns>
        public Chart AddChart(double left = 0.0d, double top = 0.0d)
        {
            Microsoft.Office.Interop.Excel.XlChartType use_ChartType =
                Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;
            object tempshapes = ExcelUtilityMethod.GetProperty(m_workSheet, "Shapes");

            object tempShape = ExcelUtilityMethod.UseMethod(tempshapes, "AddChart",
                new object[] { use_ChartType, left, top });

            object tempChart = ExcelUtilityMethod.GetProperty(tempShape, "Chart");

            return new Chart(tempChart);
        }
    }
}


