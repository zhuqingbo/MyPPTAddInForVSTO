using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MyPPTAddIn.MyUtils
{
    public static class Win32Helper
    {
        /// <summary>
        /// 该函数检索指定坐标点的像素的RGB颜色值。
        /// </summary>
        /// <param name="hDC">设备环境句柄。</param>
        /// <param name="nXPos">指定要检查的像素点的逻辑X轴坐标。</param>
        /// <param name="nYPos">指定要检查的像素点的逻辑Y轴坐标。</param>
        /// <returns>返回值是该象像点的RGB值。如果指定的像素点在当前剪辑区之外；那么返回值是CLR_INVALID。</returns>
        [DllImport("gdi32")]
        public static extern uint GetPixel(IntPtr hDC, int nXPos, int nYPos);

        /// <summary>
        /// 该函数检索一指定窗口的客户区域或整个屏幕的显示设备上下文环境的句柄，
        /// 以后可以在GDI函数中使用该句柄来在设备上下文环境中绘图。
        /// </summary>
        /// <param name="hWnd">设备上下文环境被检索的窗口的句柄，如果该值为NULL，GetDC则检索整个屏幕的设备上下文环境。</param>
        /// <returns>如果成功，返回指定窗口客户区的设备上下文环境；如果失败，返回值为Null。</returns>
        [DllImport("user32")]
        public static extern IntPtr GetDC(IntPtr hWnd);

        /// <summary>
        /// 该函数释放设备上下文环境（DC）供其他应用程序使用。函数的效果与设备上下文环境类型有关。
        /// 它只释放公用的和设备上下文环境，对于类或私有的则无效。
        /// </summary>
        /// <param name="hWnd">指向要释放的设备上下文环境所在的窗口的句柄。</param>
        /// <param name="hDC">指向要释放的设备上下文环境的句柄。</param>
        /// <returns>如果释放成功，则返回值为1；如果没有释放成功，则返回值为0。</returns>
        [DllImport("user32")]
        public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        /// <summary>
        /// 2014-12-5 21:43:51
        /// 获取一个RGB颜色值中的红色强度值。
        /// </summary>
        /// <param name="rgb">指定的RGB颜色值。</param>
        /// <returns></returns>
        public static byte GetRValue(uint rgb)
        {
            return (byte)rgb;
        }

        /// <summary>
        /// 2014-12-5 21:51:24
        /// 获取一个RGB颜色值中的绿色强度值。
        /// </summary>
        /// <param name="rgb">指定的RGB颜色值。</param>
        /// <returns></returns>
        public static byte GetGValue(uint rgb)
        {
            return (byte)(((ushort)(rgb)) >> 8);
        }

        /// <summary>
        /// 2014-12-5 21:52:37
        /// 获取一个RGB颜色值中的蓝色强度值。
        /// </summary>
        /// <param name="rgb">指定的RGB颜色值。</param>
        /// <returns></returns>
        public static byte GetBValue(uint rgb)
        {
            return (byte)(rgb >> 16);
        }


        #region DLL导入
        /// <summary>
        /// 用于设置窗口
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="hWndInsertAfter"></param>
        /// <param name="X"></param>
        /// <param name="Y"></param>
        /// <param name="cx"></param>
        /// <param name="cy"></param>
        /// <param name="uFlags"></param>
        /// <returns></returns>
        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern bool SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int X, int Y, int cx, int cy, int uFlags);

        /// <summary>
        /// 安装钩子
        /// </summary>
        /// <param name="idHook">钩子类型,鼠标\键盘\巨集等10几种</param>
        /// <param name="lpfn">挂钩的函数,用来处理拦截消息的函数,全局函数</param>
        /// <param name="hInstance">当前进程的句柄,
        /// 为NULL:当前进程创建的一个线程,子程位于当前进程；
        /// 为0(IntPtr.Zero):钩子子程与所有的线程关联，即为全局钩子</param>
        /// <param name="threadId">设置要挂接的线程ID.为NULL则为全局钩子</param>
        /// <returns></returns>
        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr SetWindowsHookEx(IntPtr idHook, HookProc lpfn, IntPtr pInstance, uint threadId);

        /// <summary>
        /// 卸载钩子
        /// </summary>
        /// <param name="idHook"></param>
        /// <returns></returns>
        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern bool UnhookWindowsHookEx(IntPtr pHookHandle);

        /// <summary>
        /// 传递钩子
        /// 用于把拦截的消息继续传递下去，不然其他程序可能会得不到相应的消息
        /// </summary>
        /// <param name="idHook"></param>
        /// <param name="nCode"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern int CallNextHookEx(IntPtr pHookHandle, int nCode, IntPtr wParam, IntPtr lParam);

        /// <summary>
        /// 转换当前按键信息
        /// </summary>
        /// <param name="uVirtKey"></param>
        /// <param name="uScanCode"></param>
        /// <param name="lpbKeyState"></param>
        /// <param name="lpwTransKey"></param>
        /// <param name="fuState"></param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern int ToAscii(UInt32 uVirtKey, UInt32 uScanCode, byte[] lpbKeyState, byte[] lpwTransKey, UInt32 fuState);

        /// <summary>
        /// 获取按键状态
        /// </summary>
        /// <param name="pbKeyState"></param>
        /// <returns>非0表示成功</returns>
        [DllImport("user32.dll")]
        public static extern int GetKeyboardState(byte[] pbKeyState);

        [DllImport("user32.dll")]
        public static extern short GetKeyStates(int vKey);

        /// <summary>
        /// 获取当前线程Id
        /// </summary>
        /// <returns></returns>
        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        /// <summary>
        /// 截屏位置
        /// </summary>
        /// <param name="hdcDest">目标设备的句柄 </param>
        /// <param name="nXDest">目标对象的左上角的X坐标</param>
        /// <param name="nYDest">目标对象的左上角的Y坐标</param>
        /// <param name="nWidth">目标对象的矩形的宽度</param>
        /// <param name="nHeight">目标对象的矩形的高度 </param>
        /// <param name="hdcSrc">源设备的句柄 </param>
        /// <param name="nXSrc">源对象的左上角的X坐标 </param>
        /// <param name="nYSrc">源对象的左上角的Y坐标 </param>
        /// <param name="dwRop">光栅的操作值 </param>
        /// <returns></returns>
        [DllImportAttribute("gdi32.dll")]
        private static extern bool BitBlt(IntPtr hdcDest, int nXDest, int nYDest, int nWidth, int nHeight,
            IntPtr hdcSrc, int nXSrc, int nYSrc, System.Int32 dwRop);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="lpszDriver">驱动名称</param>
        /// <param name="lpszDevice">设备名称</param>
        /// <param name="lpszOutput">无用，可以设定为"NULL"</param>
        /// <param name="lpInitData">任意的打印机数据</param>
        /// <returns></returns>
        [DllImportAttribute("gdi32.dll")]
        private static extern IntPtr CreateDC(string lpszDriver, string lpszDevice, string lpszOutput, IntPtr lpInitData);
        #endregion DLL导入

        /// <summary>
        /// 钩子委托声明
        /// </summary>
        /// <param name="nCode"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        public delegate int HookProc(int nCode, IntPtr wParam, IntPtr lParam);
    }
}
