namespace WindowsFormsApp1
{
    partial class Form2
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
            this.comboBox_select_file = new System.Windows.Forms.ComboBox();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.btn_select_file_Click = new System.Windows.Forms.Button();
            this.btn_select_sheet_Click = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // comboBox_select_file
            // 
            this.comboBox_select_file.FormattingEnabled = true;
            this.comboBox_select_file.Location = new System.Drawing.Point(143, 41);
            this.comboBox_select_file.Name = "comboBox_select_file";
            this.comboBox_select_file.Size = new System.Drawing.Size(121, 20);
            this.comboBox_select_file.TabIndex = 0;
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(143, 109);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(121, 20);
            this.comboBox2.TabIndex = 1;
            // 
            // btn_select_file_Click
            // 
            this.btn_select_file_Click.Location = new System.Drawing.Point(384, 41);
            this.btn_select_file_Click.Name = "btn_select_file_Click";
            this.btn_select_file_Click.Size = new System.Drawing.Size(75, 23);
            this.btn_select_file_Click.TabIndex = 2;
            this.btn_select_file_Click.Text = "btn_select_sheet_Click";
            this.btn_select_file_Click.UseVisualStyleBackColor = true;
            // 
            // btn_select_sheet_Click
            // 
            this.btn_select_sheet_Click.Location = new System.Drawing.Point(384, 109);
            this.btn_select_sheet_Click.Name = "btn_select_sheet_Click";
            this.btn_select_sheet_Click.Size = new System.Drawing.Size(75, 23);
            this.btn_select_sheet_Click.TabIndex = 3;
            this.btn_select_sheet_Click.Text = "btn_select_sheet_Click";
            this.btn_select_sheet_Click.UseVisualStyleBackColor = true;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btn_select_sheet_Click);
            this.Controls.Add(this.btn_select_file_Click);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.comboBox_select_file);
            this.Name = "Form2";
            this.Text = "Form2";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBox_select_file;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Button btn_select_file_Click;
        private System.Windows.Forms.Button btn_select_sheet_Click;
    }
}