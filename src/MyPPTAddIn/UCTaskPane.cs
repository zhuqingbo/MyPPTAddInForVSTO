using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;

namespace MyPPTAddIn
{
    public partial class UCTaskPane : UserControl
    {
        public UCTaskPane()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ////Microsoft.Office.Interop.PowerPoint.Shape textbox;
            ////Microsoft.Office.Interop.PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;//获取当前应用程序的所有PPT文档
            ////for (int i = 1; i <= slides.Count; i++)//遍历该文档集合,添加文本框
            ////{
            ////    Microsoft.Office.Interop.PowerPoint.Slide slide = slides[i];

            ////    //slide.Shapes.Count
            ////    for (int ix = 1; ix < slide.Shapes.Count + 1; ix++)
            ////    {
            ////        Shape item = slide.Shapes[ix];
            ////        //MessageBox.Show(item.TextEffect.Text);
            ////        item.Select();
            ////    }

            ////    //textbox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 50, 100, 600, 50);//向当前PPT添加文本框


            ////    //textbox.TextFrame.TextRange.Text = "测试文本";//设置文本框的内容
            ////　　  //textbox.TextFrame.TextRange.Font.Size = 48;//设置文本字体大小
            ////    //textbox.TextFrame.TextRange.Font.Color.RGB = Color.DarkViolet.ToArgb();//设置文本颜色
            ////}

            GetSelectCont();
            //GetSelectCont();


            
        }
        
        private void GetSelectCont()
        {
            Selection sec = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            try
            {
                for (int i = 1; i < sec.ShapeRange.Count + 1; i++)
                {
                    MessageBox.Show(sec.ShapeRange[i].TextEffect.Text);
                }
            }
            catch
            {
            }
            //string word = sec.TextRange.Text;
            ////return word.Trim();
            //MessageBox.Show(word);
        }
        //private void RGB()
        //{
        //    Image I; Graphics G;
        //    I = new Bitmap(Screen.PrimaryScreen.Bounds.Size.Width, Screen.PrimaryScreen.Bounds.Size.Height);
        //    int i = Control.MousePosition.X;
        //    int u = Control.MousePosition.Y;//这是获取鼠标坐标
        //    Bitmap jpg = new Bitmap(I);
        //    Color pixelColor = jpg.GetPixel(i, u); //获取一点的信息（颜色)
        //}

    }
}
