using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyPPTAddIn.MyUtils
{
    public class MouseHook
    {
        #region 定义变量
        //定义钩子处理函数
        private Win32Helper.HookProc MouseHookProcedure;
        //定义钩子句柄
        private IntPtr hHook = IntPtr.Zero;
        //定义鼠标事件
        public event MouseEventHandler OnMouseActivity;
        #endregion

        /// <summary>
        /// 安装鼠标钩子
        /// </summary>
        public void InstallHook()
        {
            if (hHook == IntPtr.Zero)
            {
                uint id = Win32Helper.GetCurrentThreadId();
                this.MouseHookProcedure = new Win32Helper.HookProc(this.MouseHookProc);
                //这里挂节钩子
                hHook = Win32Helper.SetWindowsHookEx((IntPtr)HookHelper.WH_Codes.WH_MOUSE_LL, MouseHookProcedure, IntPtr.Zero, id);
            }
        }

        /// <summary>
        /// 卸载鼠标钩子
        /// </summary>
        public void UnInstallHook()
        {
            bool isSuccess = false;
            if (this.hHook != IntPtr.Zero)
            {
                isSuccess = Win32Helper.UnhookWindowsHookEx(hHook);
                this.hHook = IntPtr.Zero;
            }
            if (isSuccess)
            {
                MessageBox.Show("卸载成功！");
            }
            else
            {
                MessageBox.Show("卸载失败！");
            }
        }

        /// <summary>
        /// 鼠标钩子处理函数
        /// </summary>
        /// <param name="nCode"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        private int MouseHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < (int)HookHelper.WH_Codes.HC_ACTION)
            {
                return Win32Helper.CallNextHookEx(hHook, nCode, wParam, lParam);
            }

            if (OnMouseActivity != null)
            {
                //Marshall the data from callback.
                HookHelper.MouseHookStruct mouseHookStruct = (HookHelper.MouseHookStruct)Marshal.PtrToStructure(lParam, typeof(HookHelper.MouseHookStruct));
                MouseButtons button = MouseButtons.None;
                short mouseDelta = 0;
                switch ((int)wParam)
                {
                    case (int)HookHelper.WM_MOUSE.WM_LBUTTONDOWN:
                        //case WM_LBUTTONUP: 
                        //case WM_LBUTTONDBLCLK:
                        button = MouseButtons.Left;
                        break;
                    case (int)HookHelper.WM_MOUSE.WM_RBUTTONDOWN:
                        //case WM_RBUTTONUP: 
                        //case WM_RBUTTONDBLCLK: 
                        button = MouseButtons.Right;
                        break;
                    case (int)HookHelper.WM_MOUSE.WM_MOUSEWHEEL:
                        //button = MouseButtons.Middle;//滚动轮
                        //(value >> 16) & 0xffff; retrieves the high-order word from the given 32-bit value
                        mouseDelta = (short)((mouseHookStruct.MouseData >> 16) & 0xffff);
                        break;
                }

                int clickCount = 0;//点击数
                if (button != MouseButtons.None)
                {
                    if (wParam == (IntPtr)HookHelper.WM_MOUSE.WM_LBUTTONDBLCLK || wParam == (IntPtr)HookHelper.WM_MOUSE.WM_RBUTTONDBLCLK)
                    {
                        clickCount = 2;//双击
                    }
                    else
                    {
                        clickCount = 1;//单击
                    }
                }

                //鼠标事件传递数据
                MouseEventArgs e = new MouseEventArgs(button, clickCount, mouseHookStruct.Point.X, mouseHookStruct.Point.Y, mouseDelta);

                //重写事件
                OnMouseActivity(this, e);
            }

            return Win32Helper.CallNextHookEx(hHook, nCode, wParam, lParam);
        }
    }
}
