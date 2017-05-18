using MyPPTAddIn.MyUtils;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace MyPPTAddIn.dlgFrm
{
    public partial class frmColorPickup : Form
    {
        public event EventHandler SetColor;
        private KeyHook _keyHook = new KeyHook();

        public frmColorPickup()
        {
            InitializeComponent();
        }
        public Color PickupColor = Color.Black;

        private uint ParseRGB(Color color)
        {
            return (uint)(((uint)color.B << 16) | (ushort)(((ushort)color.G << 8) | color.R));
        }
        
        private void frmColorPickup_Load(object sender, EventArgs e)
        {
            _hdc = Win32Helper.GetDC(_hWnd);
            timerSelectColor.Start();
            _keyHook.OnKeyDown += _keyHook_OnKeyDown; ;
            _keyHook.InstallHook();
        }

        private void _keyHook_OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Escape)
            {
                timerSelectColor.Stop();
                _keyHook.UnInstallHook();
                
                if (PickupColor.ToArgb() == Color.Black.ToArgb())
                {
                    lblDescribe.ForeColor = Color.LightGreen;
                }
                lblDescribe.Text="颜色选择完成";
            }
        }
        

        private void frmColorPickup_FormClosing(object sender, FormClosingEventArgs e)
        {
            timerSelectColor.Stop();
            _keyHook.UnInstallHook();
        }

        private void timerSelectColor_Tick(object sender, EventArgs e)
        {
            this.Activate();
            //timerSelectColor.Enabled = false;
            Point p = MousePosition;
            uint color = Win32Helper.GetPixel(_hdc, p.X, p.Y);
            byte r = Win32Helper.GetRValue(color);
            byte g = Win32Helper.GetGValue(color);
            byte b = Win32Helper.GetBValue(color);
            panelColor.BackColor = Color.FromArgb(r, g, b);
            PickupColor = panelColor.BackColor;
            txtRGB.Text = r + "," + g + "," + b;
            txtHex.Text= r.ToString("X2") + g.ToString("X2") + b.ToString("X2");
            #region 截图
            int w = 60;
            int h = 20;
            //////// 新建一个和屏幕大小相同的图片
            //////Bitmap CatchBmp = new Bitmap(w * 2, h * 2);

            //////// 创建一个画板，让我们可以在画板上画图
            //////// 这个画板也就是和屏幕大小一样大的图片
            //////// 我们可以通过Graphics这个类在这个空白图片上画图
            Point mPoint = new Point(Control.MousePosition.X, Control.MousePosition.Y);


            txtPoint.Text = Control.MousePosition.X+","+ Control.MousePosition.Y;
            Point startPoint = new Point(mPoint.X - w + 1, mPoint.Y - h );
            
            // 新建一个和屏幕大小相同的图片
            Bitmap CatchBmp = new Bitmap(w * 2, h * 2);

            // 创建一个画板，让我们可以在画板上画图
            // 这个画板也就是和屏幕大小一样大的图片
            // 我们可以通过Graphics这个类在这个空白图片上画图
            Graphics gh = Graphics.FromImage(CatchBmp);

            // 把屏幕图片拷贝到我们创建的空白图片 CatchBmp中
            gh.CopyFromScreen(startPoint, new Point(0, 0), CatchBmp.Size);//new Size(Screen.AllScreens[0].Bounds.Width, Screen.AllScreens[0].Bounds.Height));
            gh.Dispose();
            panelImg.BackgroundImage = CatchBmp;
            #endregion
            //timerSelectColor.Enabled = true;
        }
        
        /// <summary>
        /// 显示设备上下文环境的句柄。
        /// </summary>
        private IntPtr _hdc = IntPtr.Zero;

        /// <summary>
        /// 指向窗口的句柄。
        /// </summary>
        private readonly IntPtr _hWnd = IntPtr.Zero;

        private void btnSave_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Yes;
            if (SetColor != null)
            {
                SetColor(PickupColor, e);
            }
            this.Close();
        }
        
    }
}
