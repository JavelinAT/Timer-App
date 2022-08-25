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
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.label_Team_Information = new System.Windows.Forms.Label();
            this.textBox_Team_Information = new System.Windows.Forms.TextBox();
            this.label_Round_Time = new System.Windows.Forms.Label();
            this.textBox_Round_Time = new System.Windows.Forms.TextBox();
            this.splitContainer3 = new System.Windows.Forms.SplitContainer();
            this.splitContainer4 = new System.Windows.Forms.SplitContainer();
            this.textBox_Round = new System.Windows.Forms.TextBox();
            this.label_Round = new System.Windows.Forms.Label();
            this.textBox_Total_Time = new System.Windows.Forms.TextBox();
            this.label_Total_Time = new System.Windows.Forms.Label();
            this.splitContainer5 = new System.Windows.Forms.SplitContainer();
            this.label_Score = new System.Windows.Forms.Label();
            this.textBox_Score = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer3)).BeginInit();
            this.splitContainer3.Panel1.SuspendLayout();
            this.splitContainer3.Panel2.SuspendLayout();
            this.splitContainer3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer4)).BeginInit();
            this.splitContainer4.Panel1.SuspendLayout();
            this.splitContainer4.Panel2.SuspendLayout();
            this.splitContainer4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer5)).BeginInit();
            this.splitContainer5.Panel1.SuspendLayout();
            this.splitContainer5.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.splitContainer2);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.splitContainer3);
            this.splitContainer1.Size = new System.Drawing.Size(1904, 852);
            this.splitContainer1.SplitterDistance = 170;
            this.splitContainer1.TabIndex = 0;
            // 
            // splitContainer2
            // 
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.label_Team_Information);
            this.splitContainer2.Panel1.Controls.Add(this.textBox_Team_Information);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.label_Round_Time);
            this.splitContainer2.Panel2.Controls.Add(this.textBox_Round_Time);
            this.splitContainer2.Size = new System.Drawing.Size(1904, 170);
            this.splitContainer2.SplitterDistance = 516;
            this.splitContainer2.TabIndex = 0;
            // 
            // label_Team_Information
            // 
            this.label_Team_Information.AutoSize = true;
            this.label_Team_Information.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label_Team_Information.Font = new System.Drawing.Font("新細明體", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_Team_Information.Location = new System.Drawing.Point(0, 0);
            this.label_Team_Information.Name = "label_Team_Information";
            this.label_Team_Information.Size = new System.Drawing.Size(197, 27);
            this.label_Team_Information.TabIndex = 0;
            this.label_Team_Information.Text = "Team information";
            // 
            // textBox_Team_Information
            // 
            this.textBox_Team_Information.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.textBox_Team_Information.Enabled = false;
            this.textBox_Team_Information.Font = new System.Drawing.Font("Times New Roman", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_Team_Information.Location = new System.Drawing.Point(0, 52);
            this.textBox_Team_Information.Multiline = true;
            this.textBox_Team_Information.Name = "textBox_Team_Information";
            this.textBox_Team_Information.ReadOnly = true;
            this.textBox_Team_Information.Size = new System.Drawing.Size(516, 118);
            this.textBox_Team_Information.TabIndex = 2;
            this.textBox_Team_Information.TabStop = false;
            this.textBox_Team_Information.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label_Round_Time
            // 
            this.label_Round_Time.AutoSize = true;
            this.label_Round_Time.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label_Round_Time.Font = new System.Drawing.Font("新細明體", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_Round_Time.Location = new System.Drawing.Point(0, 0);
            this.label_Round_Time.Name = "label_Round_Time";
            this.label_Round_Time.Size = new System.Drawing.Size(142, 27);
            this.label_Round_Time.TabIndex = 0;
            this.label_Round_Time.Text = "Round Time";
            // 
            // textBox_Round_Time
            // 
            this.textBox_Round_Time.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.textBox_Round_Time.Enabled = false;
            this.textBox_Round_Time.Font = new System.Drawing.Font("Times New Roman", 72F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_Round_Time.ForeColor = System.Drawing.SystemColors.ControlText;
            this.textBox_Round_Time.Location = new System.Drawing.Point(0, 52);
            this.textBox_Round_Time.Name = "textBox_Round_Time";
            this.textBox_Round_Time.ReadOnly = true;
            this.textBox_Round_Time.Size = new System.Drawing.Size(1384, 118);
            this.textBox_Round_Time.TabIndex = 1;
            this.textBox_Round_Time.TabStop = false;
            this.textBox_Round_Time.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // splitContainer3
            // 
            this.splitContainer3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer3.Location = new System.Drawing.Point(0, 0);
            this.splitContainer3.Name = "splitContainer3";
            // 
            // splitContainer3.Panel1
            // 
            this.splitContainer3.Panel1.Controls.Add(this.splitContainer4);
            // 
            // splitContainer3.Panel2
            // 
            this.splitContainer3.Panel2.Controls.Add(this.splitContainer5);
            this.splitContainer3.Size = new System.Drawing.Size(1904, 678);
            this.splitContainer3.SplitterDistance = 482;
            this.splitContainer3.TabIndex = 0;
            // 
            // splitContainer4
            // 
            this.splitContainer4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer4.Location = new System.Drawing.Point(0, 0);
            this.splitContainer4.Name = "splitContainer4";
            this.splitContainer4.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer4.Panel1
            // 
            this.splitContainer4.Panel1.Controls.Add(this.textBox_Round);
            this.splitContainer4.Panel1.Controls.Add(this.label_Round);
            // 
            // splitContainer4.Panel2
            // 
            this.splitContainer4.Panel2.Controls.Add(this.textBox_Total_Time);
            this.splitContainer4.Panel2.Controls.Add(this.label_Total_Time);
            this.splitContainer4.Size = new System.Drawing.Size(482, 678);
            this.splitContainer4.SplitterDistance = 304;
            this.splitContainer4.TabIndex = 0;
            // 
            // textBox_Round
            // 
            this.textBox_Round.Enabled = false;
            this.textBox_Round.Font = new System.Drawing.Font("新細明體", 72F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.textBox_Round.Location = new System.Drawing.Point(0, 120);
            this.textBox_Round.Name = "textBox_Round";
            this.textBox_Round.ReadOnly = true;
            this.textBox_Round.Size = new System.Drawing.Size(481, 123);
            this.textBox_Round.TabIndex = 3;
            this.textBox_Round.TabStop = false;
            this.textBox_Round.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label_Round
            // 
            this.label_Round.AutoSize = true;
            this.label_Round.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label_Round.Font = new System.Drawing.Font("新細明體", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_Round.Location = new System.Drawing.Point(0, 0);
            this.label_Round.Name = "label_Round";
            this.label_Round.Size = new System.Drawing.Size(81, 27);
            this.label_Round.TabIndex = 0;
            this.label_Round.Text = "Round";
            // 
            // textBox_Total_Time
            // 
            this.textBox_Total_Time.Enabled = false;
            this.textBox_Total_Time.Font = new System.Drawing.Font("新細明體", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.textBox_Total_Time.Location = new System.Drawing.Point(1, 165);
            this.textBox_Total_Time.Name = "textBox_Total_Time";
            this.textBox_Total_Time.ReadOnly = true;
            this.textBox_Total_Time.Size = new System.Drawing.Size(481, 49);
            this.textBox_Total_Time.TabIndex = 4;
            this.textBox_Total_Time.TabStop = false;
            this.textBox_Total_Time.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label_Total_Time
            // 
            this.label_Total_Time.AutoSize = true;
            this.label_Total_Time.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label_Total_Time.Font = new System.Drawing.Font("新細明體", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_Total_Time.Location = new System.Drawing.Point(0, 0);
            this.label_Total_Time.Name = "label_Total_Time";
            this.label_Total_Time.Size = new System.Drawing.Size(118, 27);
            this.label_Total_Time.TabIndex = 0;
            this.label_Total_Time.Text = "Total time";
            // 
            // splitContainer5
            // 
            this.splitContainer5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer5.Location = new System.Drawing.Point(0, 0);
            this.splitContainer5.Name = "splitContainer5";
            // 
            // splitContainer5.Panel1
            // 
            this.splitContainer5.Panel1.Controls.Add(this.label_Score);
            this.splitContainer5.Panel1.Controls.Add(this.textBox_Score);
            this.splitContainer5.Size = new System.Drawing.Size(1418, 678);
            this.splitContainer5.SplitterDistance = 678;
            this.splitContainer5.TabIndex = 0;
            // 
            // label_Score
            // 
            this.label_Score.AutoSize = true;
            this.label_Score.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label_Score.Font = new System.Drawing.Font("新細明體", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_Score.Location = new System.Drawing.Point(0, 0);
            this.label_Score.Name = "label_Score";
            this.label_Score.Size = new System.Drawing.Size(69, 27);
            this.label_Score.TabIndex = 0;
            this.label_Score.Text = "Score";
            // 
            // textBox_Score
            // 
            this.textBox_Score.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox_Score.Enabled = false;
            this.textBox_Score.Font = new System.Drawing.Font("新細明體", 72F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.textBox_Score.Location = new System.Drawing.Point(0, 0);
            this.textBox_Score.Multiline = true;
            this.textBox_Score.Name = "textBox_Score";
            this.textBox_Score.ReadOnly = true;
            this.textBox_Score.Size = new System.Drawing.Size(678, 678);
            this.textBox_Score.TabIndex = 4;
            this.textBox_Score.TabStop = false;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(1904, 852);
            this.Controls.Add(this.splitContainer1);
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form2";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel1.PerformLayout();
            this.splitContainer2.Panel2.ResumeLayout(false);
            this.splitContainer2.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            this.splitContainer3.Panel1.ResumeLayout(false);
            this.splitContainer3.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer3)).EndInit();
            this.splitContainer3.ResumeLayout(false);
            this.splitContainer4.Panel1.ResumeLayout(false);
            this.splitContainer4.Panel1.PerformLayout();
            this.splitContainer4.Panel2.ResumeLayout(false);
            this.splitContainer4.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer4)).EndInit();
            this.splitContainer4.ResumeLayout(false);
            this.splitContainer5.Panel1.ResumeLayout(false);
            this.splitContainer5.Panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer5)).EndInit();
            this.splitContainer5.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.Label label_Team_Information;
        private System.Windows.Forms.Label label_Round_Time;
        private System.Windows.Forms.SplitContainer splitContainer3;
        private System.Windows.Forms.SplitContainer splitContainer4;
        private System.Windows.Forms.Label label_Round;
        private System.Windows.Forms.Label label_Total_Time;
        private System.Windows.Forms.SplitContainer splitContainer5;
        private System.Windows.Forms.Label label_Score;
        private System.Windows.Forms.TextBox textBox_Round_Time;
        private System.Windows.Forms.TextBox textBox_Team_Information;
        private System.Windows.Forms.TextBox textBox_Round;
        private System.Windows.Forms.TextBox textBox_Total_Time;
        private System.Windows.Forms.TextBox textBox_Score;
    }
}