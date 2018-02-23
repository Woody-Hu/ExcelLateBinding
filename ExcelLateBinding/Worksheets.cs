using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLateBinding
{
    /// <summary>
    /// Worksheets
    /// </summary>
    public class Worksheets
    {
        /// <summary>
        /// WorkSheets组件
        /// </summary>
        private object m_workSheets;

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="thisobject"></param>
        internal Worksheets(object thisobject)
        {
            m_workSheets = thisobject;
        }

        /// <summary>
        /// 获得其中一个工作表
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public Worksheet Get_Item(int index)
        {
            object thisObject = ExcelUtilityMethod.GetProperty(m_workSheets, "Item", new object[] { index });

            return new Worksheet(thisObject);
        }

        /// <summary>
        /// 获得当前工作表的数量
        /// </summary>
        public int Count
        {
            get
            {
                return (int)ExcelUtilityMethod.GetProperty(m_workSheets, "Count");
            }
        }
    }
}