using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using System.Drawing;

namespace MyPPTAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ////注意：在ThisAddIn_Startup中调用对应的方法,
            //AddText();//添加自定义文本信息
            //AddTaskPane();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        #region Test
        //添加自定义文本信息
        private void AddText()
        {
            //事件委托绑定(对PPT中添加代码控制文本)
            this.Application.PresentationNewSlide += new PowerPoint.EApplication_PresentationNewSlideEventHandler(Application_PresentationNewSlide);
        }

        //此为Application对象的PresentationNewSlide事件
        //功能：当用户将新幻灯片添加到活动演示文稿时，此事件处理程序会将文本框添加到新幻灯片的顶部，然后向文本框中添加一些文本。
        void Application_PresentationNewSlide(PowerPoint.Slide Sld)
        {
            //这里的Application表示 PowerPoint 的当前实例。
            //这里的参数Sld,表示新幻灯片的Slide对象。
            PowerPoint.Shape textBox = Sld.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            textBox.TextFrame.TextRange.InsertAfter("This Text Was Added By Using Code!");
        }

        /// <summary>
        /// 添加自定义的Task Pane
        /// </summary>
        private void AddTaskPane()
        {
            //自定义用户控件名称(即自定任务窗格类)
            UCTaskPane taskPane = new UCTaskPane();
            //将用户控件添加到 CustomTaskPaneCollection集合中
            CustomTaskPane myCustomTaskPane = this.CustomTaskPanes.Add(taskPane, "My Task Pane");
            myCustomTaskPane.Width = 200;//设置自定义任务窗格的宽度
            myCustomTaskPane.Visible = true;//设置其可见
        }
        #endregion
        
        private void AddTextBox(PowerPoint.Slide slide, string txtContent)
        {

            
            PowerPoint.Shape textbox;
            textbox = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 100, 600, 50);//向当前PPT添加文本框             
            textbox.TextFrame.TextRange.Text = txtContent;//设置文本框的内容             
            textbox.TextFrame.TextRange.Font.Size = 48;//设置文本字体大小             
            textbox.TextFrame.TextRange.Font.Color.RGB = Color.DarkViolet.ToArgb();//设置文本颜色        
        }
        private void AddPicture(PowerPoint.Slide slide, PowerPoint.Shape shape, string filePath)
        {
            PowerPoint.Shape pic;
            pic = slide.Shapes.AddPicture(filePath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, shape.Left, shape.Top, shape.Width, shape.Height);
            pic.Name = "图片";
            pic.Height = shape.Height;
            pic.Width = shape.Width;
        }
        private void AddTextMessage(PowerPoint.Slide slide, PowerPoint.Shape shape, string txtContent)
        {
            PowerPoint.Shape textbox; textbox = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, shape.Left, shape.Top, shape.Width, shape.Height);//向当前PPT添加文本框            
            textbox.Height = shape.Height;
            textbox.Width = shape.Width;
            textbox.TextFrame.TextRange.Text = txtContent;
            textbox.TextFrame.TextRange.Font.Color.RGB = Color.Orange.ToArgb();
        }
    }
}
