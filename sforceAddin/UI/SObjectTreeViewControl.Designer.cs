namespace sforceAddin.UI
{
    partial class SObjectTreeViewControl
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
            this.tv_sobjs = new System.Windows.Forms.TreeView();
            this.SuspendLayout();
            // 
            // tv_sobjs
            // 
            this.tv_sobjs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tv_sobjs.Location = new System.Drawing.Point(0, 0);
            this.tv_sobjs.Name = "tv_sobjs";
            this.tv_sobjs.Size = new System.Drawing.Size(150, 150);
            this.tv_sobjs.TabIndex = 0;
            // 
            // SObjectTreeViewControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tv_sobjs);
            this.Name = "SObjectTreeViewControl";
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.TreeView tv_sobjs;
    }
}
