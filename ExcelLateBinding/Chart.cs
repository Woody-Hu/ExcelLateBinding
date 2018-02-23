using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLateBinding
{
    /// <summary>
    /// Chart组件
    /// </summary>
    public class Chart
    {
        /// <summary>
        /// 内部Chart引用
        /// </summary>
        private object m_chart;

        /// <summary>
        /// 构造Chart组件
        /// </summary>
        /// <param name="thisObject"></param>
        internal Chart(object thisObject)
        {
            m_chart = thisObject;
        }

        /// <summary>
        /// 判断Chart是否有标题
        /// </summary>
        public bool HasTitle
        {
            set
            {
                ExcelUtilityMethod.SetProperty(m_chart, "HasTitle", new object[] { value });
            }
            get
            {
                return (bool)ExcelUtilityMethod.GetProperty(m_chart, "HasTitle");
            }
        }

        /// <summary>
        /// 为Chart设置数据源
        /// </summary>
        /// <param name="input"></param>
        public void SetSourceData(Range input)
        {
            Microsoft.Office.Interop.Excel.XlRowCol useType =
                Microsoft.Office.Interop.Excel.XlRowCol.xlColumns;
            ExcelUtilityMethod.UseMethod(m_chart, "SetSourceData",
                new object[] { input.ThisRangeObject, useType });
        }

        /// <summary>
        /// 设置图例
        /// </summary>
        public void ApplyDataLabels()
        {
            ExcelUtilityMethod.UseMethod(m_chart, "ApplyDataLabels", null);
        }

        /// <summary>
        /// 获取/更改图表标题
        /// </summary>
        public string ChartTitle
        {
            set
            {
                object tempObject = ExcelUtilityMethod.GetProperty(m_chart, "ChartTitle");
                ExcelUtilityMethod.SetProperty(tempObject, "Text", new object[] { value });
            }
            get
            {
                object tempObject = ExcelUtilityMethod.GetProperty(m_chart, "ChartTitle");
                return (string)ExcelUtilityMethod.GetProperty(tempObject, "Text");
            }
        }
    }
}
