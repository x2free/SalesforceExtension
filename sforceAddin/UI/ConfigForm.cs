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
    public partial class ConfigForm : Form
    {
        public Func<string, bool> APIVersionChnagedHandler;

        private bool isAPIVersionChanged = false;
        public ConfigForm()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterParent;

            this.Init();
        }

        private void Init()
        {
            this.textBox_APIVersion.Text = string.Format("{0}.0", Auth.AuthUtil.apiVersion);
        }

        private void textBox_APIVersion_TextChanged(object sender, EventArgs e)
        {
            double version;
            if (!double.TryParse(this.textBox_APIVersion.Text, out version))
            {
                // MessageBox.Show("Invalid verion number", "sforce Addin", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.textBox_APIVersion.BackColor = Color.Red;
                this.isAPIVersionChanged = false;
            }
            else
            {
                this.isAPIVersionChanged = true;
                this.textBox_APIVersion.BackColor = default(Color);
            }
        }

        private void btn_Ok_Click(object sender, EventArgs e)
        {
            if (this.isAPIVersionChanged)
            {
                if (APIVersionChnagedHandler != null)
                {
                    APIVersionChnagedHandler(this.textBox_APIVersion.Text);
                }

                this.isAPIVersionChanged = false;
            }

            this.Close();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            this.Init();
        }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox_APIVersion_Validating(object sender, CancelEventArgs e)
        {
        }
    }
}
