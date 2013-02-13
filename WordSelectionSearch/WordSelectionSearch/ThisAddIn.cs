using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools;
using System.Windows.Forms;

namespace WordSelectionSearch
{
    public partial class ThisAddIn
    {
        private CustomTaskPane taskpane = null;//word自定义面板对象
        private WebbrowserPane webbrowserPane = null;//用户控件
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
           //添加浏览器控件
            webbrowserPane = new WebbrowserPane();
            taskpane = this.CustomTaskPanes.Add(webbrowserPane, "网页浏览器");
            taskpane.Width = 400;
            taskpane.Visible = false;

            //添加菜单按钮
            RemoveAddedMenuItems();
            AddRightClickMenuItems();
            //添加鼠标右键的事件
            this.Application.WindowBeforeRightClick += new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(Application_WindowBeforeRightClick);
        }
        //删除新增鼠标右键菜单按钮
        void RemoveAddedMenuItems()
        {
            Office.CommandBarButton DelBTN = null;
            DelBTN = (Office.CommandBarButton)Application.CommandBars["Text"].FindControl(Office.MsoControlType.msoControlButton, missing, "AddedSearchBTN", missing, false);

            do 
            {
                if (DelBTN!=null)
                {
                    DelBTN.Delete(true);
                }
                DelBTN = (Office.CommandBarButton)Application.CommandBars["Text"].FindControl(Office.MsoControlType.msoControlButton, missing, "AddedSearchBTN", missing, false);
            } while (DelBTN!=null);
        }

        //新增鼠标右键菜单按钮
        void AddRightClickMenuItems()
        {
            Office.CommandBarButton AddBTN = null;
            AddBTN = (Office.CommandBarButton)Application.CommandBars["Text"].Controls.Add(Office.MsoControlType.msoControlButton,missing,missing,1, true);
            AddBTN.Tag = "AddedSearchBTN";
            AddBTN.Enabled = false;
            AddBTN.Caption = "搜索";
            AddBTN.Click += new Office._CommandBarButtonEvents_ClickEventHandler(AddBTN_Click);
        }

        void AddBTN_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            this.taskpane.Visible = true;
            this.webbrowserPane.search(this.Application.Selection.Range.Text);
        }

        void Application_WindowBeforeRightClick(Word.Selection Sel, ref bool Cancel)
        {
            //如果sel的文本不为空，则激活按钮
            if (!string.IsNullOrWhiteSpace(Sel.Range.Text))
            {
                Office.CommandBarButton searchBTN = (Office.CommandBarButton)Application.CommandBars["Text"].FindControl(Office.MsoControlType.msoControlButton, missing, "AddedSearchBTN", missing, false);
                searchBTN.Enabled = true;
            }
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            RemoveAddedMenuItems();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
