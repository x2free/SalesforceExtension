using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace sforceAddin.UI
{
    public partial class QueryForm : Form
    {
        public Func<string, bool> FilterChangedHandler;

        public QueryForm()
        {
            InitializeComponent();
        }

        public string GetFilter()
        {
            return this.textBox_filter.Text;
        }

        public void SetSelect(string strSelect)
        {
            this.textBox_select.Text = strSelect;
        }

        public void SetFilter(string strFilter)
        {
            this.textBox_filter.Text = strFilter;
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            if (FilterChangedHandler != null)
            {
                FilterChangedHandler(this.textBox_filter.Text);
            }

            this.Close();
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
