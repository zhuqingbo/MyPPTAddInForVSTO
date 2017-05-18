using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyPPTAddIn.MyUtils
{
    public class KeyHook
    {
        //定义钩子处理函数
        private Win32Helper.HookProc KeyboardProcDelegate = null;
        //定义钩子句柄
        private IntPtr khook;
        //定义键盘事件
        public event KeyEventHandler OnKeyDown;

        /// <summary>
        /// 安装钩子
        /// </summary>
        public void InstallHook()
        {
            if (khook == IntPtr.Zero)
            {
                uint id = Win32Helper.GetCurrentThreadId();
                this.KeyboardProcDelegate = new Win32Helper.HookProc(this.KeyboardProc);
                khook = Win32Helper.SetWindowsHookEx((IntPtr)HookHelper.WH_Codes.WH_KEYBOARD_LL, this.KeyboardProcDelegate, IntPtr.Zero, id);
            }
        }

        /// <summary>
        /// 卸载钩子
        /// </summary>
        public void UnInstallHook()
        {
            bool isSuccess = false;
            if (khook != IntPtr.Zero)
            {
                isSuccess = Win32Helper.UnhookWindowsHookEx(khook);
                this.khook = IntPtr.Zero;
            }
            //if (isSuccess)
            //{
            //    MessageBox.Show("卸载成功！");
            //}
            //else
            //{
            //    MessageBox.Show("卸载失败！");
            //}
        }

        /// <summary>
        /// 键盘钩子处理函数
        /// </summary>
        /// <param name="code"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        private int KeyboardProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode < (int)HookHelper.WH_Codes.HC_ACTION)
                {
                    return Win32Helper.CallNextHookEx((IntPtr)khook, nCode, wParam, lParam);
                }

                //HookHelper.KeyboardHookStruct keyHookStruct = (HookHelper.KeyboardHookStruct)Marshal.PtrToStructure(lParam, typeof(HookHelper.KeyboardHookStruct));
                
                #region
                //if ((int)wParam == (int)Keys.C && ((int)lParam & (int)Keys.ControlKey) != 0 ||//Ctrl+C
                //    (int)wParam == (int)Keys.X && ((int)lParam & (int)Keys.ControlKey) != 0)//Ctrl+V
                //{
                //    MessageBox.Show("C||V");
                //}

                //if (lParam.ToInt32() > 0)//捕获键盘按下
                //{
                //    Keys keys = (Keys)wParam;
                //    MessageBox.Show("键盘按下");
                //}
                //if (lParam.ToInt32() < 0)//捕获键盘抬起
                //{
                //    MessageBox.Show("键盘抬起");
                //}
                /**************** 
                 //全局键盘钩子判断是否按下键 
                 wParam = = 0x100 // 键盘按下 
                 wParam = = 0x101 // 键盘抬起 
                ****************/
                //return 0;//如果返回1,则结束消息,消息截止,不再传递。如果返回0或调用CallNextHookEx函数,消息出这个钩子继续往下传递。
                #endregion  
                bool handled = false;
                if (lParam.ToInt32() > 0)//捕获键盘按下
                {
                    Keys keyData = (Keys)wParam; //(Keys)keyHookStruct.VKCode;
                    KeyEventArgs e = new KeyEventArgs(keyData);
                    OnKeyDown.Invoke(this, e);
                    handled = e.Handled;
                }

                //if (this.OnKeyDown != null && (wParam == (IntPtr)HookHelper.WM_KEYBOARD.WM_KEYDOWN || wParam == (IntPtr)HookHelper.WM_KEYBOARD.WM_SYSKEYDOWN))
                //{
                //    Keys keyData = (Keys)keyHookStruct.VKCode;
                //    KeyEventArgs e = new KeyEventArgs(keyData);
                //    OnKeyDown.Invoke(this, e);
                //    handled = e.Handled;
                //}

                if (handled)
                    return -1;

                return Win32Helper.CallNextHookEx(khook, nCode, wParam, lParam);
            }
            catch
            {

            }

            return Win32Helper.CallNextHookEx(khook, nCode, wParam, lParam);
        }
    }
}
