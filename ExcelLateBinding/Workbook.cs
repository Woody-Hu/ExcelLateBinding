using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLateBinding
{ /// <summary>
  /// WorkBook组件
  /// </summary>
    public class Workbook
    {
        /// <summary>
        /// WorkBook组件
        /// </summary>
        object m_workBook;

        /// <summary>
        /// Path
        /// </summary>
        public string Path
        {
            get
            {
                object returnValue = ExcelUtilityMethod.GetProperty(m_workBook, "Path");
                if (returnValue is string)
                {
                    return (string)returnValue;
                }
                else
                {
                    return string.Empty;
                }
            }
        }

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="thisObject"></param>
        internal Workbook(object thisObject)
        {
            m_workBook = thisObject;
        }

        /// <summary>
        /// 获取所有的工作表
        /// </summary>
        public Worksheets Worksheets
        {
            get
            {
                return new Worksheets(ExcelUtilityMethod.GetProperty(m_workBook, "Worksheets"));
            }

        }

        /// <summary>
        /// 保存
        /// </summary>
        public void Save()
        {
            ExcelUtilityMethod.UseMethod(m_workBook, "Save", null);
        }

        /// <summary>
        /// 另存为
        /// </summary>
        /// <param name="strpath"></param>
        public void SaveAs(string strpath)
        {
            ExcelUtilityMethod.UseMethod(m_workBook, "SaveAs", new object[] { strpath });
        }

        /// <summary>
        /// 添加Chart
        /// </summary>
        /// <returns></returns>
        public Chart AddChart()
        {
            Microsoft.Office.Interop.Excel.XlChartType use_ChartType =
               Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;
            Microsoft.Office.Interop.Excel.XlChartLocation use_XlLocation =
                Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAutomatic;
            object chartsObject = ExcelUtilityMethod.GetProperty(m_workBook, "Charts");
            object addedChartObject = ExcelUtilityMethod.UseMethod(chartsObject, "Add",
                new object[] { Type.Missing, Type.Missing, 1 });
            //图表形式
            ExcelUtilityMethod.SetProperty(addedChartObject, "ChartType", new object[] { use_ChartType });
            //图表位置
            ExcelUtilityMethod.UseMethod(addedChartObject, "Location", new object[] { use_XlLocation });
            return new Chart(addedChartObject);
        }

        /// <summary>
        /// 全名称
        /// </summary>
        /// <returns></returns>
        public string FullName()
        {
            object returnValue = ExcelUtilityMethod.UseMethod(m_workBook, "FullName", null);
            if (returnValue is string)
            {
                return (string)returnValue;
            }
            else //返回有误时（不是字符串格式）返回空字符串
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// 关闭工作簿
        /// </summary>
        public void Close()
        {
            ExcelUtilityMethod.UseMethod(m_workBook, "Close", new object[] { null, null, null });
        }
    }
}
