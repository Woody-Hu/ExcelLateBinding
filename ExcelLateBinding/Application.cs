using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLateBinding
{
    /// <summary>
    /// Applicantion封装
    /// </summary>
    public class Application
    {
        #region 进程用

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        //函数原型；DWORD GetWindowThreadProcessld(HWND hwnd，LPDWORD lpdwProcessld);
        //参数：hWnd:窗口句柄
        //参数：lpdwProcessld:接收进程标识的32位值的地址。如果这个参数不为NULL，GetWindwThreadProcessld将进程标识拷贝到这个32位值中，否则不拷贝
        //返回值：返回值为创建窗口的线程标识。
        [DllImport("kernel32.dll")]
        static extern int OpenProcess(int dwDesiredAccess, bool bInheritHandle, int dwProcessId);
        //函数原型：HANDLE OpenProcess(DWORD dwDesiredAccess,BOOL bInheritHandle,DWORD dwProcessId);
        //参数：dwDesiredAccess：访问权限。
        //参数：bInheritHandle：继承标志。
        //参数：dwProcessId：进程ID。
        const int PROCESS_ALL_ACCESS = 0x1F0FFF;
        const int PROCESS_VM_READ = 0x0010;
        const int PROCESS_VM_WRITE = 0x0020;
        //定义句柄变量
        IntPtr hwnd;
        //定义进程ID变量
        int pid = 0;
        #endregion

        /// <summary>
        /// 应用程序组件
        /// </summary>
        object latApplication;

        /// <summary>
        /// WorkBooks
        /// </summary>
        public Workbooks Workbooks
        {
            get
            {
                object workBooks = ExcelUtilityMethod.GetProperty(latApplication, "Workbooks");
                return new Workbooks(workBooks);
            }
        }

        /// <summary>
        /// 应用程序
        /// </summary>
        /// <exception cref="ArgumentException">没有安装Excel时抛出</exception>
        public Application()
        {
            try
            {
                latApplication = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                //获取Excel App的句柄
                hwnd = new IntPtr(this.Hwnd);
                //通过Windows API获取Excel进程ID
                GetWindowThreadProcessId(hwnd, out pid);
            }
            catch
            {
                throw new ArgumentException("没有安装Microsoft Office Excel。");
            }

            //属性名
            string[] propertyNames =
                    new string[] { "DisplayAlerts", "AlertBeforeOverwriting"
                    ,"Visible","Interactive","UserControl","AskToUpdateLinks",
                    "AutoFormatAsYouTypeReplaceHyperlinks","DisplayClipboardWindow"
                    ,"DisplayDocumentActionTaskPane","DisplayDocumentInformationPanel"
                    ,"DisplayExcel4Menus","EnableCheckFileExtensions","EnableLargeOperationAlert"
                    ,"MergeInstances","MouseAvailable","ScreenUpdating"};

            foreach (var propertyName in propertyNames)
            {
                try
                {
                    ExcelUtilityMethod.SetProperty(latApplication, propertyName, new object[] { false });
                }
                catch (Exception) // 当设置属性出现异常时
                {
                    continue; //跳过
                }
            }
        }

        /// <summary>
        /// 退出
        /// </summary>
        public void Quit()
        {
            ExcelUtilityMethod.UseMethod(latApplication, "Quit", null);
            try
            {
                Process tempProcess = Process.GetProcessById(pid);
                tempProcess.Kill();
            }
            catch
            {
                ;
            }
        }

        /// <summary>
        /// 主窗体句柄
        /// </summary>
        /// <returns></returns>
        private int Hwnd
        {
            get
            {
                return (int)ExcelUtilityMethod.GetProperty(latApplication, "Hwnd", null);
            }
        }
    }
}
