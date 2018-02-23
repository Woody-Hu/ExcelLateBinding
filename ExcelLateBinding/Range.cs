using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLateBinding
{
    /// <summary>
    /// Range组件
    /// </summary>
    public class Range
    {
        /// <summary>
        /// Range组件
        /// </summary>
        object m_range;

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="thisObject"></param>
        internal Range(object thisObject)
        {
            m_range = thisObject;
        }

        /// <summary>
        /// 获取Range对象
        /// </summary>
        internal object ThisRangeObject
        {
            get
            {
                return m_range;
            }
        }

        /// <summary>
        /// 获取Range的左位置
        /// </summary>
        public dynamic Left
        {
            get
            {
                return ExcelUtilityMethod.GetProperty(m_range, "Left");
            }
        }

        /// <summary>
        /// 获取Range的上位置
        /// </summary>
        public dynamic Top
        {
            get
            {
                return ExcelUtilityMethod.GetProperty(m_range, "Top");
            }
        }

        /// <summary>
        /// Range的列宽
        /// </summary>
        public dynamic ColumnWidth
        {
            set
            {
                ExcelUtilityMethod.SetProperty(m_range, "ColumnWidth", new object[] { value });
            }
            get
            {
                return ExcelUtilityMethod.GetProperty(m_range, "ColumnWidth");
            }
        }

        /// <summary>
        /// Range的行高
        /// </summary>
        public dynamic RowHeight
        {
            set
            {
                ExcelUtilityMethod.SetProperty(m_range, "RowHeight", new object[] { value });
            }
            get
            {
                return ExcelUtilityMethod.GetProperty(m_range, "RowHeight");
            }
        }

        /// <summary>
        /// Range的值
        /// </summary>
        public dynamic Value
        {
            set
            {
                ExcelUtilityMethod.SetProperty(m_range, "Value", new object[] { value });
            }
            get
            {
                return ExcelUtilityMethod.GetProperty(m_range, "Value");
            }
        }

        /// <summary>
        /// Range的值
        /// </summary>
        public dynamic Value2
        {
            set
            {
                ExcelUtilityMethod.SetProperty(m_range, "Value2", new object[] { value });
            }
            get
            {
                return ExcelUtilityMethod.GetProperty(m_range, "Value2");
            }
        }

        /// <summary>
        /// 为Range设置文字
        /// </summary>
        public dynamic Text
        {
            get
            {
                return ExcelUtilityMethod.GetProperty(m_range, "Text");
            }
        }

        /// <summary>
        /// 批量输入数据
        /// </summary>
        /// <param name="input"></param>
        public void SetValue2(object[,] input)
        {
            ExcelUtilityMethod.SetProperty(m_range, "Value2", new object[] { input });
        }

        /// <summary>
        /// 获得终止行
        /// </summary>
        /// <returns></returns>
        public int EndRow()
        {
            object tempEnd = ExcelUtilityMethod.GetProperty(m_range, "End",
                new object[] { Microsoft.Office.Interop.Excel.XlDirection.xlDown });
            return (int)ExcelUtilityMethod.GetProperty(tempEnd, "Row");
        }

        /// <summary>
        /// 自动调整
        /// </summary>
        public void AutoFit()
        {
            object temp = ExcelUtilityMethod.GetProperty(m_range, "EntireColumn");
            ExcelUtilityMethod.UseMethod(temp, "AutoFit", null);
            return;
        }

    }
}
