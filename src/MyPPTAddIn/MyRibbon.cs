using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Drawing;
using MyPPTAddIn.dlgFrm;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace MyPPTAddIn
{
    public partial class MyRibbon
    {
        

        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnColorPanel_Click(object sender, RibbonControlEventArgs e)
        {
            DialogResult dlg=  colorDialog.ShowDialog();
            if (dlg == DialogResult.OK)
            {
                Color sColor = colorDialog.Color;
                SetSelectionForeColor(sColor);
            }
        }
        private void SetSelectionForeColor(Color color)
        {
            Microsoft.Office.Interop.PowerPoint.Selection sec = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            try
            {

                for (int i = 1; i < sec.ShapeRange.Count + 1; i++)
                {
                    sec.ShapeRange[i].TextFrame.TextRange.Font.Color.RGB = (int)ParseRGB(color);
                }
            }
            catch
            {
            }
        }
        private uint ParseRGB(Color color)
        {
            return (uint)(((uint)color.B << 16) | (ushort)(((ushort)color.G << 8) | color.R));
        }

        private Color RGB()
        {
            Image I;
            //Graphics G;
            I = new Bitmap(Screen.PrimaryScreen.Bounds.Size.Width, Screen.PrimaryScreen.Bounds.Size.Height);
            int i = Control.MousePosition.X;
            int u = Control.MousePosition.Y;//这是获取鼠标坐标
            Bitmap jpg = new Bitmap(I);
            Color pixelColor = jpg.GetPixel(i, u); //获取一点的信息（颜色)
            return pixelColor;
        }

        private void btnColorPickup_Click(object sender, RibbonControlEventArgs e)
        {
            frmColorPickup frm = new frmColorPickup();
            frm.SetColor += Frm_SetColor;
            frm.Show();
           
        }

        private void Frm_SetColor(object sender, EventArgs e)
        {
            SetSelectionForeColor((Color)sender);
        }

        private void btnAddFooter_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.PowerPoint.Shape textbox;
            Microsoft.Office.Interop.PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;//获取当前应用程序的所有PPT文档
            for (int i = 1; i <= slides.Count; i++)//遍历该文档集合,添加文本框
            {
                Microsoft.Office.Interop.PowerPoint.Slide slide = slides[i];
                bool hasfooter = false;
                ////slide.Shapes.Count
                for (int ix = 1; ix < slide.Shapes.Count + 1; ix++)
                {
                    Microsoft.Office.Interop.PowerPoint.Shape item = slide.Shapes[ix];
                    ////MessageBox.Show(item.TextEffect.Text);
                    //item.Select();
                    if (item.Name == "myFooter")
                    {
                        hasfooter =true;
                        break;
                    }
                }
                if (!hasfooter)
                {

                    //var shape= slide.Shapes.AddCallout(MsoCalloutType.msoCalloutTwo, 10, 10, 100, 200);
                    //slide.Shapes.AddChart2(1, XlChartType.xl3DArea, 100, 100, 600, 200, true);//折线图
                    //slide.Shapes.AddConnector(MsoConnectorType.msoConnectorElbow, 100, 100, 400, 400);//曲线
                    textbox = slide.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 200, 50);//向当前PPT添加文本框
                    textbox.Name = "myFooter";
                    DateTime.Now.ToString("yy.MM.dd");
                    textbox.TextFrame.TextRange.Text ="我是页脚";// "XX."+ DateTime.Now.ToString("yyyy-MM-dd"); //设置文本框的内容
                    textbox.TextFrame.TextRange.Font.Size = 12;//设置文本字体大小
                    //textbox.Fill.BackColor.RGB = (int)ParseRGB(Color.Yellow);
                    //textbox.TextFrame.TextRange.Font.Color.RGB = Color.DarkViolet.ToArgb();//设置文本颜色
                    textbox.Top = 505;
                }
            }
        }

    }
}
