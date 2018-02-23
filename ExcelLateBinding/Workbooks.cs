using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLateBinding
{
    /// <summary>
    /// WorkBooks组件
    /// </summary>
    public class Workbooks
    {
        /// <summary>
        /// WorkBooks组件
        /// </summary>
        object workbooks;

        internal Workbooks(object thisWorkBooks)
        {
            workbooks = thisWorkBooks;
        }

        /// <summary>
        /// 打开一个文档
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public Workbook Open(string path)
        {
            object thisObject;

            try
            {
                thisObject = ExcelUtilityMethod.UseMethod(workbooks, "Open", new object[] { path, false, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value });
            }
            catch (Exception)
            {
                thisObject = ExcelUtilityMethod.UseMethod(workbooks, "Open", new object[] { path, false });
            }


            return new Workbook(thisObject);
        }

        /// <summary>
        /// 新建一个文档
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public Workbook Add()
        {
            object thisObject = ExcelUtilityMethod.UseMethod(workbooks, "Add", null);
            return new Workbook(thisObject);
        }

        /// <summary>
        /// 关闭此文档
        /// </summary>
        public void Close()
        {
            ExcelUtilityMethod.UseMethod(workbooks, "Close", null);
        }

    }
}
