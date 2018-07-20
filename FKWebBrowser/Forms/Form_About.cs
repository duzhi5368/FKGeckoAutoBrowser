using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;

namespace FKWebBrowser.Forms
{
    public partial class Form_About : Form
    {
        public Form_About()
        {
            InitializeComponent();
        }

        private void Form_About_Load(object sender, EventArgs e)
        {
            // 窗口标题修改
            this.Text = Language.LanguageRes.WndTitle_About;

            // 控件名修改
            labelAbout.Text = Language.LanguageRes.Desc_About;
        }
    }
}
