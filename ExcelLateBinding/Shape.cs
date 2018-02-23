using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLateBinding
{
    /// <summary>
    /// Shape组件
    /// </summary>
    public class Shape
    {
        /// <summary>
        /// Shape封装
        /// </summary>
        private object m_thisShape;

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="shape"></param>
        internal Shape(object shape)
        {
            m_thisShape = shape;
        }

        /// <summary>
        /// 锁定比例
        /// </summary>
        public Microsoft.Office.Core.MsoTriState LockAspectRatio
        {
            set
            {
                ExcelUtilityMethod.SetProperty(m_thisShape, "LockAspectRatio", new object[] { value });
            }
            get
            {
                return (Microsoft.Office.Core.MsoTriState)ExcelUtilityMethod.GetProperty(m_thisShape, "LockAspectRatio");
            }
        }

        /// <summary>
        /// 宽度
        /// </summary>
        public float Width
        {
            set
            {
                ExcelUtilityMethod.SetProperty(m_thisShape, "Width", new object[] { value });
            }
            get
            {
                return (float)ExcelUtilityMethod.GetProperty(m_thisShape, "Width");
            }
        }

        /// <summary>
        /// 高度
        /// </summary>
        public float Height
        {
            set
            {
                ExcelUtilityMethod.SetProperty(m_thisShape, "Height", new object[] { value });
            }
            get
            {
                return (float)ExcelUtilityMethod.GetProperty(m_thisShape, "Height");
            }
        }

        /// <summary>
        /// 将图片拷贝到剪切板
        /// </summary>
        public void CopyPicture()
        {
            ExcelUtilityMethod.UseMethod(m_thisShape, "CopyPicture",
                new object[]{
                    Microsoft.Office.Interop.Excel.XlPictureAppearance.xlScreen,
                     Microsoft.Office.Interop.Excel.XlCopyPictureFormat.xlBitmap});
        }



    }
}
