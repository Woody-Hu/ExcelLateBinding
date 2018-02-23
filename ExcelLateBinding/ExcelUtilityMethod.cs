using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

/// <summary>
/// 此空间公开了一些操作Excel的公共方法
/// 后期绑定模式
/// </summary>
namespace ExcelLateBinding
{
    /// <summary>
    /// 包内Excel后期绑定公共方法
    /// </summary>
    internal static class ExcelUtilityMethod
    {
        /// <summary>
        /// 通过后期加载设置一个属性
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="propertyName"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        internal static void SetProperty(object obj, string propertyName, object[] values)
        {
            try
            {
                obj.GetType().InvokeMember(propertyName, BindingFlags.SetProperty,
                null, obj, values);
            }
            catch (Exception ex)
            {
                Type thisType = obj.GetType();
                PropertyInfo[] propArray = thisType.GetProperties();
                List<string> propNames = new List<string>();

                foreach (var item in propArray)
                {
                    propNames.Add(item.Name);
                }
                throw ex;
            }

        }

        /// <summary>
        /// 获取一个属性
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        internal static object GetProperty(object obj, string propertyName)
        {
            return obj.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, null, obj, null);
        }

        /// <summary>
        /// 获取一个属性
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        internal static object GetProperty(object obj, string propertyName, object[] values)
        {
            return obj.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, null, obj, values);
        }

        /// <summary>
        /// 使用方法
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="MethodName"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        internal static object UseMethod(object obj, string MethodName, object[] values)
        {
            try
            {
                object tempObject = obj.GetType().InvokeMember(MethodName, BindingFlags.InvokeMethod, null, obj, values);
                return tempObject;
            }
            catch (Exception)
            {
                object tempApplication = GetProperty(obj, "Application");
                try
                {
                    SetProperty(tempApplication, "Visible", new object[] { true });
                    object tempObject = obj.GetType().InvokeMember(MethodName, BindingFlags.InvokeMethod, null, obj, values);
                    SetProperty(tempObject, "Visible", new object[] { false });
                    return tempObject;
                }
                catch (Exception) //再发生异常的话，暂时什么都不做
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// 字符串转换为Int 在Excel中
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static int TitleToNumber(string s)
        {
            string sUP = s.ToUpper();
            int returnValue = 0;
            int stringLength = s.Length;
            char a = 'A';

            for (int i = 0; i < stringLength; i++)
            {
                int tempValue = (sUP[i] - a) + 1;
                tempValue = tempValue * ((int)System.Math.Pow((double)26, (double)(stringLength - i - 1)));
                returnValue = returnValue + tempValue;
            }
            return returnValue;
        }

        /// <summary>
        /// int转换为字符串 在Excel中
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string ConvertToTitle(int n)
        {
            List<int> lstNumber = new List<int>();

            int s = n;

            while (s > 26)
            {
                if (0 == s % 26)
                {
                    lstNumber.Add(26);
                    s = s - 26;
                }
                else
                {
                    lstNumber.Add(s % 26);
                    s = s - s % 26;
                }
                s = s / 26;

            }

            lstNumber.Add(s);


            lstNumber.Reverse();
            StringBuilder tempsb = new StringBuilder();
            foreach (var item in lstNumber)
            {
                tempsb.Append((char)(item - 1 + 'A'));
            }

            return tempsb.ToString();

        }
    }
}