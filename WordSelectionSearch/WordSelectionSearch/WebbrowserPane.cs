using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WordSelectionSearch
{
    public partial class WebbrowserPane : UserControl
    {
        public WebbrowserPane()
        {
            InitializeComponent();
        }
        public void search(string keyWord)
        {
            if (keyWord!="")
            {
                this.webBrowser.Navigate("www.baidu.com/s?wd="+keyWord);
                this.webBrowser.Show();
            }
        }
    }
}
