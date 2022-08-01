namespace WindowsFormsApp1
{
    partial class FrontPage
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrontPage));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage_Main = new System.Windows.Forms.TabPage();
            this.splitContainer4 = new System.Windows.Forms.SplitContainer();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.button_Command_Ready = new System.Windows.Forms.Button();
            this.button_Command_Fail = new System.Windows.Forms.Button();
            this.button_Command_3 = new System.Windows.Forms.Button();
            this.button_Command_4 = new System.Windows.Forms.Button();
            this.button_Command_Restart = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.tabPage_Setting = new System.Windows.Forms.TabPage();
            this.splitContainer3 = new System.Windows.Forms.SplitContainer();
            this.label_ComState = new System.Windows.Forms.Label();
            this.PorSelector = new System.Windows.Forms.ComboBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.textBoxSend = new System.Windows.Forms.TextBox();
            this.button_Sand = new System.Windows.Forms.Button();
            this.textBoxReceive = new System.Windows.Forms.TextBox();
            this.button_Clear = new System.Windows.Forms.Button();
            this.tabPage_Excel = new System.Windows.Forms.TabPage();
            this.splitContainer5 = new System.Windows.Forms.SplitContainer();
            this.label_excel_2 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
=========
>>>>>>>>> Temporary merge branch 2
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label_excel_1 = new System.Windows.Forms.Label();
            this.splitContainer6 = new System.Windows.Forms.SplitContainer();
            this.button_excel_1 = new System.Windows.Forms.Button();
            this.button_excel_2 = new System.Windows.Forms.Button();
            this.button_excel_5 = new System.Windows.Forms.Button();
            this.button_excel_3 = new System.Windows.Forms.Button();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
<<<<<<<<< Temporary merge branch 1
=========
            this.splitContainer5 = new System.Windows.Forms.SplitContainer();
            this.splitContainer6 = new System.Windows.Forms.SplitContainer();
            this.label_excel_2 = new System.Windows.Forms.Label();
>>>>>>>>> Temporary merge branch 2
            this.tabControl1.SuspendLayout();
            this.tabPage_Main.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer4)).BeginInit();
            this.splitContainer4.Panel1.SuspendLayout();
            this.splitContainer4.Panel2.SuspendLayout();
            this.splitContainer4.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.tabPage_Setting.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer3)).BeginInit();
            this.splitContainer3.Panel1.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            this.tabPage_Excel.SuspendLayout();
<<<<<<<<< Temporary merge branch 1
            this.panel1.SuspendLayout();
=========
>>>>>>>>> Temporary merge branch 2
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
<<<<<<<<< Temporary merge branch 1
=========
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer5)).BeginInit();
            this.splitContainer5.Panel1.SuspendLayout();
            this.splitContainer5.Panel2.SuspendLayout();
            this.splitContainer5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer6)).BeginInit();
            this.splitContainer6.Panel1.SuspendLayout();
            this.splitContainer6.Panel2.SuspendLayout();
            this.splitContainer6.SuspendLayout();
>>>>>>>>> Temporary merge branch 2
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage_Main);
            this.tabControl1.Controls.Add(this.tabPage_Setting);
            this.tabControl1.Controls.Add(this.tabPage_Excel);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(984, 442);
            this.tabControl1.TabIndex = 9;
            // 
            // tabPage_Main
            // 
            this.tabPage_Main.Controls.Add(this.splitContainer4);
            this.tabPage_Main.Location = new System.Drawing.Point(4, 28);
            this.tabPage_Main.Margin = new System.Windows.Forms.Padding(0, 1, 0, 1);
            this.tabPage_Main.Name = "tabPage_Main";
            this.tabPage_Main.Padding = new System.Windows.Forms.Padding(0, 1, 0, 1);
            this.tabPage_Main.Size = new System.Drawing.Size(976, 410);
            this.tabPage_Main.TabIndex = 0;
            this.tabPage_Main.Text = "Command";
            this.tabPage_Main.UseVisualStyleBackColor = true;
            // 
            // splitContainer4
            // 
            this.splitContainer4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer4.IsSplitterFixed = true;
            this.splitContainer4.Location = new System.Drawing.Point(0, 1);
            this.splitContainer4.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.splitContainer4.Name = "splitContainer4";
            // 
            // splitContainer4.Panel1
            // 
            this.splitContainer4.Panel1.Controls.Add(this.tableLayoutPanel1);
            // 
            // splitContainer4.Panel2
            // 
            this.splitContainer4.Panel2.Controls.Add(this.label1);
            this.splitContainer4.Size = new System.Drawing.Size(976, 408);
            this.splitContainer4.SplitterDistance = 147;
            this.splitContainer4.SplitterWidth = 2;
            this.splitContainer4.TabIndex = 0;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.button_Command_Ready, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.button_Command_Fail, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.button_Command_3, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.button_Command_4, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.button_Command_Restart, 0, 4);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 5;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(147, 408);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // button_Command_Ready
            // 
            this.button_Command_Ready.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.button_Command_Ready.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_Command_Ready.FlatAppearance.BorderSize = 0;
            this.button_Command_Ready.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Command_Ready.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_Command_Ready.Location = new System.Drawing.Point(1, 2);
            this.button_Command_Ready.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.button_Command_Ready.Name = "button_Command_Ready";
            this.button_Command_Ready.Size = new System.Drawing.Size(145, 77);
            this.button_Command_Ready.TabIndex = 0;
            this.button_Command_Ready.Text = "Ready";
            this.button_Command_Ready.UseVisualStyleBackColor = false;
            this.button_Command_Ready.Click += new System.EventHandler(this.button_Command_Ready_Click);
            // 
            // button_Command_Fail
            // 
            this.button_Command_Fail.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.button_Command_Fail.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_Command_Fail.FlatAppearance.BorderSize = 0;
            this.button_Command_Fail.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Command_Fail.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_Command_Fail.ForeColor = System.Drawing.Color.Black;
            this.button_Command_Fail.Location = new System.Drawing.Point(1, 83);
            this.button_Command_Fail.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.button_Command_Fail.Name = "button_Command_Fail";
            this.button_Command_Fail.Size = new System.Drawing.Size(145, 77);
            this.button_Command_Fail.TabIndex = 1;
            this.button_Command_Fail.Text = "Fail";
            this.button_Command_Fail.UseVisualStyleBackColor = false;
            this.button_Command_Fail.Click += new System.EventHandler(this.button_Command_Fail_Click);
            // 
            // button_Command_3
            // 
            this.button_Command_3.BackColor = System.Drawing.Color.Silver;
            this.button_Command_3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_Command_3.FlatAppearance.BorderSize = 0;
            this.button_Command_3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Command_3.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_Command_3.Location = new System.Drawing.Point(1, 164);
            this.button_Command_3.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.button_Command_3.Name = "button_Command_3";
            this.button_Command_3.Size = new System.Drawing.Size(145, 77);
            this.button_Command_3.TabIndex = 2;
            this.button_Command_3.Text = "button3";
            this.button_Command_3.UseVisualStyleBackColor = false;
            this.button_Command_3.Click += new System.EventHandler(this.button_Command_3_Click);
            // 
            // button_Command_4
            // 
            this.button_Command_4.BackColor = System.Drawing.Color.Silver;
            this.button_Command_4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_Command_4.FlatAppearance.BorderSize = 0;
            this.button_Command_4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Command_4.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_Command_4.Location = new System.Drawing.Point(1, 245);
            this.button_Command_4.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.button_Command_4.Name = "button_Command_4";
            this.button_Command_4.Size = new System.Drawing.Size(145, 77);
            this.button_Command_4.TabIndex = 3;
            this.button_Command_4.Text = "button4";
            this.button_Command_4.UseVisualStyleBackColor = false;
            this.button_Command_4.Click += new System.EventHandler(this.button_Command_4_Click);
            // 
            // button_Command_Restart
            // 
            this.button_Command_Restart.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.button_Command_Restart.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_Command_Restart.FlatAppearance.BorderSize = 0;
            this.button_Command_Restart.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Command_Restart.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_Command_Restart.ForeColor = System.Drawing.SystemColors.Control;
            this.button_Command_Restart.Location = new System.Drawing.Point(1, 326);
            this.button_Command_Restart.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.button_Command_Restart.Name = "button_Command_Restart";
            this.button_Command_Restart.Size = new System.Drawing.Size(145, 80);
            this.button_Command_Restart.TabIndex = 4;
            this.button_Command_Restart.Text = "Restart";
            this.button_Command_Restart.UseVisualStyleBackColor = false;
            this.button_Command_Restart.Click += new System.EventHandler(this.button_Command_Restart_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Font = new System.Drawing.Font("Times New Roman", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 33);
            this.label1.TabIndex = 1;
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tabPage_Setting
            // 
            this.tabPage_Setting.Controls.Add(this.splitContainer3);
            this.tabPage_Setting.Location = new System.Drawing.Point(4, 28);
            this.tabPage_Setting.Margin = new System.Windows.Forms.Padding(0, 1, 0, 1);
            this.tabPage_Setting.Name = "tabPage_Setting";
            this.tabPage_Setting.Padding = new System.Windows.Forms.Padding(0, 1, 0, 1);
            this.tabPage_Setting.Size = new System.Drawing.Size(976, 410);
            this.tabPage_Setting.TabIndex = 1;
            this.tabPage_Setting.Text = "Setting";
            this.tabPage_Setting.UseVisualStyleBackColor = true;
            // 
            // splitContainer3
            // 
            this.splitContainer3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer3.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer3.IsSplitterFixed = true;
            this.splitContainer3.Location = new System.Drawing.Point(0, 1);
            this.splitContainer3.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.splitContainer3.Name = "splitContainer3";
            this.splitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer3.Panel1
            // 
            this.splitContainer3.Panel1.Controls.Add(this.label_ComState);
            this.splitContainer3.Panel1.Controls.Add(this.PorSelector);
            // 
            // splitContainer3.Panel2
            // 
            this.splitContainer3.Panel2.Controls.Add(this.checkBox1);
            this.splitContainer3.Panel2.Controls.Add(this.splitContainer2);
            this.splitContainer3.Size = new System.Drawing.Size(976, 408);
            this.splitContainer3.SplitterDistance = 49;
            this.splitContainer3.SplitterWidth = 2;
            this.splitContainer3.TabIndex = 1;
            // 
            // label_ComState
            // 
            this.label_ComState.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.label_ComState.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label_ComState.CausesValidation = false;
            this.label_ComState.Cursor = System.Windows.Forms.Cursors.Hand;
            this.label_ComState.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label_ComState.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_ComState.Location = new System.Drawing.Point(0, 0);
            this.label_ComState.Name = "label_ComState";
            this.label_ComState.Size = new System.Drawing.Size(976, 25);
            this.label_ComState.TabIndex = 1;
            this.label_ComState.Text = "Click to select Com Port";
            this.label_ComState.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label_ComState.Click += new System.EventHandler(this.label_ComState_Click);
            // 
            // PorSelector
            // 
            this.PorSelector.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PorSelector.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.PorSelector.DropDownHeight = 110;
            this.PorSelector.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.PorSelector.DropDownWidth = 104;
            this.PorSelector.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.PorSelector.Font = new System.Drawing.Font("新細明體", 12F);
            this.PorSelector.FormattingEnabled = true;
            this.PorSelector.IntegralHeight = false;
            this.PorSelector.Location = new System.Drawing.Point(0, 25);
            this.PorSelector.Margin = new System.Windows.Forms.Padding(1);
            this.PorSelector.Name = "PorSelector";
            this.PorSelector.Size = new System.Drawing.Size(976, 24);
            this.PorSelector.Sorted = true;
            this.PorSelector.TabIndex = 2;
            this.PorSelector.TabStop = false;
            this.PorSelector.Visible = false;
            this.PorSelector.DropDown += new System.EventHandler(this.PorSelector_DropDown);
            this.PorSelector.SelectedIndexChanged += new System.EventHandler(this.PorSelector_SelectedIndexChanged);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(34, 18);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(97, 23);
            this.checkBox1.TabIndex = 2;
            this.checkBox1.Text = "checkBox1";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // splitContainer2
            // 
            this.splitContainer2.IsSplitterFixed = true;
            this.splitContainer2.Location = new System.Drawing.Point(134, 1);
            this.splitContainer2.Margin = new System.Windows.Forms.Padding(0, 1, 0, 1);
            this.splitContainer2.Name = "splitContainer2";
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.textBoxSend);
            this.splitContainer2.Panel1.Controls.Add(this.button_Sand);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.textBoxReceive);
            this.splitContainer2.Panel2.Controls.Add(this.button_Clear);
            this.splitContainer2.Size = new System.Drawing.Size(650, 311);
            this.splitContainer2.SplitterDistance = 309;
            this.splitContainer2.SplitterWidth = 1;
            this.splitContainer2.TabIndex = 0;
            // 
            // textBoxSend
            // 
            this.textBoxSend.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxSend.Location = new System.Drawing.Point(0, 44);
            this.textBoxSend.Multiline = true;
            this.textBoxSend.Name = "textBoxSend";
            this.textBoxSend.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxSend.Size = new System.Drawing.Size(309, 267);
            this.textBoxSend.TabIndex = 3;
            // 
            // button_Sand
            // 
            this.button_Sand.Dock = System.Windows.Forms.DockStyle.Top;
            this.button_Sand.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_Sand.Location = new System.Drawing.Point(0, 0);
            this.button_Sand.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button_Sand.Name = "button_Sand";
            this.button_Sand.Size = new System.Drawing.Size(309, 44);
            this.button_Sand.TabIndex = 0;
            this.button_Sand.Text = "Sand Text";
            this.button_Sand.UseVisualStyleBackColor = true;
            this.button_Sand.Click += new System.EventHandler(this.button_Sand_Click);
            // 
            // textBoxReceive
            // 
            this.textBoxReceive.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxReceive.Location = new System.Drawing.Point(0, 44);
            this.textBoxReceive.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.textBoxReceive.Multiline = true;
            this.textBoxReceive.Name = "textBoxReceive";
            this.textBoxReceive.ReadOnly = true;
            this.textBoxReceive.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxReceive.Size = new System.Drawing.Size(340, 267);
            this.textBoxReceive.TabIndex = 4;
            // 
            // button_Clear
            // 
            this.button_Clear.Dock = System.Windows.Forms.DockStyle.Top;
            this.button_Clear.Location = new System.Drawing.Point(0, 0);
            this.button_Clear.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.button_Clear.Name = "button_Clear";
            this.button_Clear.Size = new System.Drawing.Size(340, 44);
            this.button_Clear.TabIndex = 5;
            this.button_Clear.Text = "Clear Receive Box";
            this.button_Clear.UseCompatibleTextRendering = true;
            this.button_Clear.UseVisualStyleBackColor = true;
            this.button_Clear.Click += new System.EventHandler(this.button_Clear_Click);
            // 
            // tabPage_Excel
            // 
            this.tabPage_Excel.Controls.Add(this.splitContainer5);
            this.tabPage_Excel.Location = new System.Drawing.Point(4, 28);
            this.tabPage_Excel.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.tabPage_Excel.Name = "tabPage_Excel";
            this.tabPage_Excel.Padding = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.tabPage_Excel.Size = new System.Drawing.Size(976, 410);
            this.splitContainer5.Panel1.Controls.Add(this.label_excel_2);
            this.splitContainer5.Panel1.Controls.Add(this.label_excel_1);
<<<<<<<<< Temporary merge branch 1
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.dataGridView1);
            this.panel1.Controls.Add(this.label_excel_1);
            this.panel1.Controls.Add(this.button_excel_5);
            this.panel1.Controls.Add(this.button_excel_4);
            this.panel1.Controls.Add(this.button_excel_3);
            this.panel1.Controls.Add(this.button_excel_2);
            this.panel1.Controls.Add(this.button_excel_1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(1, 2);
            this.panel1.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(974, 405);
            this.panel1.TabIndex = 0;
=========
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(740, 365);
            this.dataGridView1.TabIndex = 7;
            this.dataGridView1.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellEnter);
>>>>>>>>> Temporary merge branch 2
            // 
            // splitContainer5.Panel2
            // 
            this.splitContainer5.Panel2.Controls.Add(this.splitContainer6);
            this.splitContainer5.Size = new System.Drawing.Size(974, 406);
            this.splitContainer5.SplitterDistance = 37;
            this.splitContainer5.TabIndex = 1;
            // 
            // label_excel_2
            // 
            this.label_excel_2.AutoSize = true;
            this.label_excel_2.Location = new System.Drawing.Point(481, 0);
            this.label_excel_2.Name = "label_excel_2";
            this.label_excel_2.Size = new System.Drawing.Size(45, 19);
            this.label_excel_2.TabIndex = 5;
            this.label_excel_2.Text = "label2";
            // 
            // label_excel_1
            // 
            this.label_excel_1.AutoSize = true;
            this.label_excel_1.Dock = System.Windows.Forms.DockStyle.Left;
            this.label_excel_1.Location = new System.Drawing.Point(0, 0);
            this.label_excel_1.Name = "label_excel_1";
            this.label_excel_1.Size = new System.Drawing.Size(36, 19);
            this.label_excel_1.TabIndex = 4;
            this.label_excel_1.Text = "Path";
            // 
            // splitContainer6
            // 
            this.splitContainer6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer6.Location = new System.Drawing.Point(0, 0);
            this.splitContainer6.Name = "splitContainer6";
            // 
            // splitContainer6.Panel1
            // 
            this.splitContainer6.Panel1.Controls.Add(this.button_excel_1);
            this.splitContainer6.Panel1.Controls.Add(this.button_excel_2);
            this.splitContainer6.Panel1.Controls.Add(this.button_excel_5);
            this.splitContainer6.Panel1.Controls.Add(this.button_excel_3);
            this.splitContainer6.Panel1.Controls.Add(this.button_excel_4);
            // 
            // splitContainer6.Panel2
            // 
            this.splitContainer6.Panel2.Controls.Add(this.dataGridView1);
            this.splitContainer6.Size = new System.Drawing.Size(974, 365);
            this.splitContainer6.SplitterDistance = 230;
            this.splitContainer6.TabIndex = 0;
            // 
            // button_excel_1
            // 
            this.button_excel_1.Location = new System.Drawing.Point(33, 16);
            this.button_excel_1.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.button_excel_1.Name = "button_excel_1";
            this.button_excel_1.Size = new System.Drawing.Size(185, 30);
            this.button_excel_1.TabIndex = 0;
            this.button_excel_1.Text = "OpenTemplate";
            this.button_excel_1.UseVisualStyleBackColor = true;
            this.button_excel_1.Click += new System.EventHandler(this.button_excel_1_Click);
            // 
            // button_excel_2
            // 
            this.button_excel_2.Location = new System.Drawing.Point(33, 94);
            this.button_excel_2.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.button_excel_2.Name = "button_excel_2";
            this.button_excel_2.Size = new System.Drawing.Size(185, 30);
            this.button_excel_2.TabIndex = 1;
            this.button_excel_2.Text = "SelectFolder";
            this.button_excel_2.UseVisualStyleBackColor = true;
            this.button_excel_2.Click += new System.EventHandler(this.button_excel_2_Click);
            // 
            // button_excel_5
            // 
            this.button_excel_5.Location = new System.Drawing.Point(33, 228);
            this.button_excel_5.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.button_excel_5.Name = "button_excel_5";
            this.button_excel_5.Size = new System.Drawing.Size(185, 30);
            this.button_excel_5.TabIndex = 5;
            this.button_excel_5.Text = "ReadFile";
            this.button_excel_5.UseVisualStyleBackColor = true;
            this.button_excel_5.Click += new System.EventHandler(this.button_excel_5_Click);
            // 
            // button_excel_3
            // 
            this.button_excel_3.Location = new System.Drawing.Point(33, 140);
            this.button_excel_3.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.button_excel_3.Name = "button_excel_3";
            this.button_excel_3.Size = new System.Drawing.Size(185, 30);
            this.button_excel_3.TabIndex = 2;
            this.button_excel_3.Text = "SaveOnExcel";
            this.button_excel_3.UseVisualStyleBackColor = true;
            this.button_excel_3.Click += new System.EventHandler(this.button_excel_3_Click);
            // 
            // button_excel_4
            // 
            this.button_excel_4.Location = new System.Drawing.Point(33, 184);
            this.button_excel_4.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.button_excel_4.Name = "button_excel_4";
            this.button_excel_4.Size = new System.Drawing.Size(185, 30);
            this.button_excel_4.TabIndex = 3;
            this.button_excel_4.Text = "OpenFile";
            this.button_excel_4.UseVisualStyleBackColor = true;
            this.button_excel_4.Click += new System.EventHandler(this.button_excel_4_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(740, 365);
            this.dataGridView1.TabIndex = 7;
            this.dataGridView1.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellEnter);
            // 
            // label_display
            // 
            this.label_display.BackColor = System.Drawing.Color.Transparent;
            this.label_display.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label_display.Font = new System.Drawing.Font("Times New Roman", 72F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_display.Location = new System.Drawing.Point(0, 0);
            this.splitContainer5.TabIndex = 1;
            // 
            // splitContainer6
            // 
            this.splitContainer6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer6.Location = new System.Drawing.Point(0, 0);
            this.splitContainer6.Name = "splitContainer6";
            // 
            // splitContainer6.Panel1
            // 
            this.splitContainer6.Panel1.Controls.Add(this.button_excel_1);
            this.splitContainer6.Panel1.Controls.Add(this.button_excel_2);
            this.splitContainer6.Panel1.Controls.Add(this.button_excel_5);
            this.splitContainer6.Panel1.Controls.Add(this.button_excel_3);
            this.splitContainer6.Panel1.Controls.Add(this.button_excel_4);
            // 
            // splitContainer6.Panel2
            // 
            this.splitContainer6.Panel2.Controls.Add(this.dataGridView1);
            this.splitContainer6.Size = new System.Drawing.Size(974, 365);
            this.splitContainer6.SplitterDistance = 230;
            this.splitContainer6.TabIndex = 0;
            // 
            // label_excel_2
            // 
            this.label_excel_2.AutoSize = true;
            this.label_excel_2.Location = new System.Drawing.Point(481, 0);
            this.label_excel_2.Name = "label_excel_2";
            this.label_excel_2.Size = new System.Drawing.Size(45, 19);
            this.label_excel_2.TabIndex = 5;
            this.label_excel_2.Text = "label2";
            // 
>>>>>>>>> Temporary merge branch 2
            // FrontPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MinimumSize = new System.Drawing.Size(884, 501);
            this.Name = "FrontPage";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Timer App";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.FrontPage_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage_Main.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.splitContainer4.Panel2.ResumeLayout(false);
            this.splitContainer4.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer4)).EndInit();
            this.splitContainer4.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tabPage_Setting.ResumeLayout(false);
            this.splitContainer3.Panel1.ResumeLayout(false);
            this.splitContainer3.Panel2.ResumeLayout(false);
            this.splitContainer3.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer3)).EndInit();
            this.splitContainer3.ResumeLayout(false);
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel1.PerformLayout();
            this.splitContainer2.Panel2.ResumeLayout(false);
            this.splitContainer2.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            this.tabPage_Excel.ResumeLayout(false);
<<<<<<<<< Temporary merge branch 1
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
=========
>>>>>>>>> Temporary merge branch 2
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
<<<<<<<<< Temporary merge branch 1
=========
            this.splitContainer5.Panel1.ResumeLayout(false);
            this.splitContainer5.Panel1.PerformLayout();
            this.splitContainer5.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer5)).EndInit();
            this.splitContainer5.ResumeLayout(false);
            this.splitContainer6.Panel1.ResumeLayout(false);
            this.splitContainer6.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer6)).EndInit();
            this.splitContainer6.ResumeLayout(false);
>>>>>>>>> Temporary merge branch 2
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage_Main;
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.Button button_Sand;
        private System.Windows.Forms.TextBox textBoxSend;
        private System.Windows.Forms.Label label_ComState;
        private System.Windows.Forms.ComboBox PorSelector;
        private System.Windows.Forms.TabPage tabPage_Setting;
        private System.Windows.Forms.Label label_display;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.SplitContainer splitContainer3;
        private System.Windows.Forms.SplitContainer splitContainer4;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button button_Command_Ready;
        private System.Windows.Forms.Button button_Command_Fail;
        private System.Windows.Forms.Button button_Command_3;
        private System.Windows.Forms.Button button_Command_4;
        private System.Windows.Forms.Button button_Command_Restart;
        private System.Windows.Forms.Button button_Clear;
        private System.Windows.Forms.TextBox textBoxReceive;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TabPage tabPage_Excel;
        private System.Windows.Forms.Button button_excel_3;
        private System.Windows.Forms.Button button_excel_2;
        private System.Windows.Forms.Button button_excel_1;
        private System.Windows.Forms.Button button_excel_4;
        private System.Windows.Forms.Label label_excel_1;
        private System.Windows.Forms.Button button_excel_5;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.SplitContainer splitContainer5;
        private System.Windows.Forms.SplitContainer splitContainer6;
        private System.Windows.Forms.Label label_excel_2;
    }
}

