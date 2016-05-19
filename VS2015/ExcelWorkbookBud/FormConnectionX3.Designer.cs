namespace ExcelWorkbookBud
{
    partial class FormConnectionX3
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.labelLoginX3 = new System.Windows.Forms.Label();
            this.labelPasswordX3 = new System.Windows.Forms.Label();
            this.labelLanguageX3 = new System.Windows.Forms.Label();
            this.textBoxLoginX3 = new System.Windows.Forms.TextBox();
            this.textBoxPasswordX3 = new System.Windows.Forms.TextBox();
            this.comboBoxLanguageX3 = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonOK
            // 
            this.buttonOK.Location = new System.Drawing.Point(41, 203);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(75, 23);
            this.buttonOK.TabIndex = 0;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(189, 203);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 1;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // labelLoginX3
            // 
            this.labelLoginX3.AutoSize = true;
            this.labelLoginX3.Location = new System.Drawing.Point(6, 38);
            this.labelLoginX3.Name = "labelLoginX3";
            this.labelLoginX3.Size = new System.Drawing.Size(49, 13);
            this.labelLoginX3.TabIndex = 2;
            this.labelLoginX3.Text = "Login X3";
            // 
            // labelPasswordX3
            // 
            this.labelPasswordX3.AutoSize = true;
            this.labelPasswordX3.Location = new System.Drawing.Point(6, 73);
            this.labelPasswordX3.Name = "labelPasswordX3";
            this.labelPasswordX3.Size = new System.Drawing.Size(69, 13);
            this.labelPasswordX3.TabIndex = 3;
            this.labelPasswordX3.Text = "Password X3";
            this.labelPasswordX3.Click += new System.EventHandler(this.label2_Click);
            // 
            // labelLanguageX3
            // 
            this.labelLanguageX3.AutoSize = true;
            this.labelLanguageX3.Location = new System.Drawing.Point(6, 111);
            this.labelLanguageX3.Name = "labelLanguageX3";
            this.labelLanguageX3.Size = new System.Drawing.Size(71, 13);
            this.labelLanguageX3.TabIndex = 4;
            this.labelLanguageX3.Text = "Language X3";
            // 
            // textBoxLoginX3
            // 
            this.textBoxLoginX3.Location = new System.Drawing.Point(111, 35);
            this.textBoxLoginX3.Name = "textBoxLoginX3";
            this.textBoxLoginX3.Size = new System.Drawing.Size(121, 20);
            this.textBoxLoginX3.TabIndex = 5;
            // 
            // textBoxPasswordX3
            // 
            this.textBoxPasswordX3.Location = new System.Drawing.Point(113, 66);
            this.textBoxPasswordX3.Name = "textBoxPasswordX3";
            this.textBoxPasswordX3.Size = new System.Drawing.Size(121, 20);
            this.textBoxPasswordX3.TabIndex = 7;
            this.textBoxPasswordX3.UseSystemPasswordChar = true;
            // 
            // comboBoxLanguageX3
            // 
            this.comboBoxLanguageX3.FormattingEnabled = true;
            this.comboBoxLanguageX3.Items.AddRange(new object[] {
            "ENG",
            "FRA"});
            this.comboBoxLanguageX3.Location = new System.Drawing.Point(113, 103);
            this.comboBoxLanguageX3.Name = "comboBoxLanguageX3";
            this.comboBoxLanguageX3.Size = new System.Drawing.Size(121, 21);
            this.comboBoxLanguageX3.TabIndex = 8;
            this.comboBoxLanguageX3.Text = "ENG";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBoxLoginX3);
            this.groupBox1.Controls.Add(this.textBoxPasswordX3);
            this.groupBox1.Controls.Add(this.labelLoginX3);
            this.groupBox1.Controls.Add(this.comboBoxLanguageX3);
            this.groupBox1.Controls.Add(this.labelLanguageX3);
            this.groupBox1.Controls.Add(this.labelPasswordX3);
            this.groupBox1.Location = new System.Drawing.Point(32, 13);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(255, 164);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Connection";
            // 
            // FormConnectionX3
            // 
            this.AcceptButton = this.buttonOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.buttonCancel;
            this.ClientSize = new System.Drawing.Size(352, 251);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormConnectionX3";
            this.ShowInTaskbar = false;
            this.Text = "X3 parameters";
            this.TopMost = true;
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Label labelLoginX3;
        private System.Windows.Forms.Label labelPasswordX3;
        private System.Windows.Forms.Label labelLanguageX3;
        internal System.Windows.Forms.TextBox textBoxLoginX3;
        internal System.Windows.Forms.TextBox textBoxPasswordX3;
        internal System.Windows.Forms.ComboBox comboBoxLanguageX3;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}