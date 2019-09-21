namespace sforceAddin.UI
{
    partial class sforceListViewControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.listview_sobjs = new System.Windows.Forms.ListView();
            this.SuspendLayout();
            // 
            // listview_sobjs
            // 
            this.listview_sobjs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listview_sobjs.FullRowSelect = true;
            this.listview_sobjs.GridLines = false;
            this.listview_sobjs.Location = new System.Drawing.Point(0, 0);
            this.listview_sobjs.Name = "listview_sobjs";
            // this.listview_sobjs.Size = new System.Drawing.Size(150, 150);
            this.listview_sobjs.TabIndex = 0;
            this.listview_sobjs.UseCompatibleStateImageBehavior = false;
            this.listview_sobjs.View = System.Windows.Forms.View.List;
            this.listview_sobjs.Scrollable = true;
            // this.listview_sobjs.ListViewItemSorter = ;
            // 
            // sforceListViewControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.listview_sobjs);
            this.Name = "sforceListViewControl";
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.ListView listview_sobjs;
    }
}
