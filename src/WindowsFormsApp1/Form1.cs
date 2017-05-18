using MyPPTAddIn.MyUtils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        KeyboardHookLib _keyboardHook = new KeyboardHookLib();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _keyboardHook.InstallHook(OnKeyPress);
        }

        public void OnKeyPress(KeyboardHookLib.HookStruct hookStruct, out bool handle)
        {
            handle = false; //预设不拦截任何键
            Keys key = (Keys)hookStruct.vkCode;
            MessageBox.Show(key.ToString());


        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            _keyboardHook.UninstallHook();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _keyboardHook.UninstallHook();
        }
    }
}
