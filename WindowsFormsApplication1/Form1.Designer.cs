namespace WindowsFormsApplication1
{
    partial class Excel_Gen
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
            this.project_id = new System.Windows.Forms.TextBox();
            this.time_point = new System.Windows.Forms.TextBox();
            this.layer = new System.Windows.Forms.TextBox();
            this.formulation = new System.Windows.Forms.TextBox();
            this.compound = new System.Windows.Forms.TextBox();
            this.replica = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.time_point_label = new System.Windows.Forms.Label();
            this.layer_label = new System.Windows.Forms.Label();
            this.formulation_label = new System.Windows.Forms.Label();
            this.compound_label = new System.Windows.Forms.Label();
            this.replica_label = new System.Windows.Forms.Label();
            this.button_start = new System.Windows.Forms.Button();
            this.study_type_2 = new System.Windows.Forms.RadioButton();
            this.study_type_1 = new System.Windows.Forms.RadioButton();
            this.table = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.volume = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.generate_table = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.button1 = new System.Windows.Forms.Button();
            this.button_load = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.api_count_box = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.table)).BeginInit();
            this.SuspendLayout();
            // 
            // project_id
            // 
            this.project_id.Location = new System.Drawing.Point(122, 58);
            this.project_id.Name = "project_id";
            this.project_id.Size = new System.Drawing.Size(100, 21);
            this.project_id.TabIndex = 1;
            this.project_id.Text = "P1";
            this.project_id.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // time_point
            // 
            this.time_point.Location = new System.Drawing.Point(122, 85);
            this.time_point.Name = "time_point";
            this.time_point.Size = new System.Drawing.Size(100, 21);
            this.time_point.TabIndex = 2;
            this.time_point.Text = "2";
            // 
            // layer
            // 
            this.layer.Location = new System.Drawing.Point(122, 116);
            this.layer.Name = "layer";
            this.layer.Size = new System.Drawing.Size(100, 21);
            this.layer.TabIndex = 3;
            this.layer.Text = "2";
            // 
            // formulation
            // 
            this.formulation.Location = new System.Drawing.Point(122, 143);
            this.formulation.Name = "formulation";
            this.formulation.Size = new System.Drawing.Size(100, 21);
            this.formulation.TabIndex = 4;
            this.formulation.Text = "2";
            // 
            // compound
            // 
            this.compound.Location = new System.Drawing.Point(122, 170);
            this.compound.Name = "compound";
            this.compound.Size = new System.Drawing.Size(100, 21);
            this.compound.TabIndex = 5;
            this.compound.Text = "2";
            // 
            // replica
            // 
            this.replica.Location = new System.Drawing.Point(122, 197);
            this.replica.Name = "replica";
            this.replica.Size = new System.Drawing.Size(100, 21);
            this.replica.TabIndex = 6;
            this.replica.Text = "2";
            this.replica.TextChanged += new System.EventHandler(this.replica_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 61);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 7;
            this.label1.Text = "Project ID";
            // 
            // time_point_label
            // 
            this.time_point_label.AutoSize = true;
            this.time_point_label.Location = new System.Drawing.Point(17, 88);
            this.time_point_label.Name = "time_point_label";
            this.time_point_label.Size = new System.Drawing.Size(77, 12);
            this.time_point_label.TabIndex = 8;
            this.time_point_label.Text = "# Time Point";
            // 
            // layer_label
            // 
            this.layer_label.AutoSize = true;
            this.layer_label.Location = new System.Drawing.Point(21, 116);
            this.layer_label.Name = "layer_label";
            this.layer_label.Size = new System.Drawing.Size(47, 12);
            this.layer_label.TabIndex = 9;
            this.layer_label.Text = "# Layer";
            // 
            // formulation_label
            // 
            this.formulation_label.AutoSize = true;
            this.formulation_label.Location = new System.Drawing.Point(19, 143);
            this.formulation_label.Name = "formulation_label";
            this.formulation_label.Size = new System.Drawing.Size(71, 12);
            this.formulation_label.TabIndex = 10;
            this.formulation_label.Text = "Formulation";
            // 
            // compound_label
            // 
            this.compound_label.AutoSize = true;
            this.compound_label.Location = new System.Drawing.Point(19, 170);
            this.compound_label.Name = "compound_label";
            this.compound_label.Size = new System.Drawing.Size(53, 12);
            this.compound_label.TabIndex = 11;
            this.compound_label.Text = "Compound";
            // 
            // replica_label
            // 
            this.replica_label.AutoSize = true;
            this.replica_label.Location = new System.Drawing.Point(21, 197);
            this.replica_label.Name = "replica_label";
            this.replica_label.Size = new System.Drawing.Size(47, 12);
            this.replica_label.TabIndex = 12;
            this.replica_label.Text = "Replica";
            // 
            // button_start
            // 
            this.button_start.Location = new System.Drawing.Point(19, 260);
            this.button_start.Name = "button_start";
            this.button_start.Size = new System.Drawing.Size(103, 23);
            this.button_start.TabIndex = 13;
            this.button_start.Text = "Generate Params";
            this.button_start.UseVisualStyleBackColor = true;
            this.button_start.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // study_type_2
            // 
            this.study_type_2.AutoSize = true;
            this.study_type_2.BackColor = System.Drawing.SystemColors.Control;
            this.study_type_2.Location = new System.Drawing.Point(122, 21);
            this.study_type_2.Name = "study_type_2";
            this.study_type_2.Size = new System.Drawing.Size(59, 16);
            this.study_type_2.TabIndex = 15;
            this.study_type_2.TabStop = true;
            this.study_type_2.Text = "Type 2";
            this.study_type_2.UseVisualStyleBackColor = false;
            this.study_type_2.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // study_type_1
            // 
            this.study_type_1.AutoSize = true;
            this.study_type_1.Checked = true;
            this.study_type_1.Location = new System.Drawing.Point(21, 21);
            this.study_type_1.Name = "study_type_1";
            this.study_type_1.Size = new System.Drawing.Size(59, 16);
            this.study_type_1.TabIndex = 16;
            this.study_type_1.TabStop = true;
            this.study_type_1.Text = "Type 1";
            this.study_type_1.UseVisualStyleBackColor = true;
            this.study_type_1.CheckedChanged += new System.EventHandler(this.radioButton2_CheckedChanged);
            // 
            // table
            // 
            this.table.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.table.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3,
            this.volume,
            this.Column5,
            this.Column6});
            this.table.Location = new System.Drawing.Point(12, 289);
            this.table.Name = "table";
            this.table.RowTemplate.Height = 23;
            this.table.Size = new System.Drawing.Size(366, 436);
            this.table.TabIndex = 17;
            this.table.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.table_CellContentClick);
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Name";
            this.Column1.Name = "Column1";
            // 
            // Column2
            // 
            this.Column2.HeaderText = "Index";
            this.Column2.Name = "Column2";
            this.Column2.Width = 50;
            // 
            // Column3
            // 
            this.Column3.HeaderText = "value";
            this.Column3.Name = "Column3";
            // 
            // volume
            // 
            this.volume.HeaderText = "Volume";
            this.volume.Name = "volume";
            // 
            // Column5
            // 
            this.Column5.HeaderText = "Dosed Before";
            this.Column5.Name = "Column5";
            // 
            // Column6
            // 
            this.Column6.HeaderText = "Dosed After";
            this.Column6.Name = "Column6";
            // 
            // generate_table
            // 
            this.generate_table.Location = new System.Drawing.Point(128, 260);
            this.generate_table.Name = "generate_table";
            this.generate_table.Size = new System.Drawing.Size(104, 23);
            this.generate_table.TabIndex = 18;
            this.generate_table.Text = "Generate Labels";
            this.generate_table.UseVisualStyleBackColor = true;
            this.generate_table.Click += new System.EventHandler(this.generate_table_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Location = new System.Drawing.Point(384, 12);
            this.tabControl1.MinimumSize = new System.Drawing.Size(800, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(800, 770);
            this.tabControl1.TabIndex = 19;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(238, 260);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(112, 23);
            this.button1.TabIndex = 20;
            this.button1.Text = "Export Label";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button_load
            // 
            this.button_load.Location = new System.Drawing.Point(229, 21);
            this.button_load.Name = "button_load";
            this.button_load.Size = new System.Drawing.Size(121, 23);
            this.button_load.TabIndex = 21;
            this.button_load.Text = "Load Input";
            this.button_load.UseVisualStyleBackColor = true;
            this.button_load.Click += new System.EventHandler(this.button2_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(229, 50);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(121, 23);
            this.button2.TabIndex = 22;
            this.button2.Text = "Export Report";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 225);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 12);
            this.label2.TabIndex = 23;
            this.label2.Text = "API count";
            // 
            // api_count_box
            // 
            this.api_count_box.Location = new System.Drawing.Point(122, 225);
            this.api_count_box.Name = "api_count_box";
            this.api_count_box.Size = new System.Drawing.Size(100, 21);
            this.api_count_box.TabIndex = 24;
            this.api_count_box.Text = "1";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(238, 231);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(112, 23);
            this.button3.TabIndex = 25;
            this.button3.Text = "Print Label";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_2);
            // 
            // Excel_Gen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1188, 753);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.api_count_box);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button_load);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.generate_table);
            this.Controls.Add(this.table);
            this.Controls.Add(this.study_type_1);
            this.Controls.Add(this.study_type_2);
            this.Controls.Add(this.button_start);
            this.Controls.Add(this.replica_label);
            this.Controls.Add(this.compound_label);
            this.Controls.Add(this.formulation_label);
            this.Controls.Add(this.layer_label);
            this.Controls.Add(this.time_point_label);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.replica);
            this.Controls.Add(this.compound);
            this.Controls.Add(this.formulation);
            this.Controls.Add(this.layer);
            this.Controls.Add(this.time_point);
            this.Controls.Add(this.project_id);
            this.Name = "Excel_Gen";
            this.Text = "Excel Gen";
            this.Load += new System.EventHandler(this.Excel_Gen_Load);
            ((System.ComponentModel.ISupportInitialize)(this.table)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox project_id;
        private System.Windows.Forms.TextBox time_point;
        private System.Windows.Forms.TextBox layer;
        private System.Windows.Forms.TextBox formulation;
        private System.Windows.Forms.TextBox compound;
        private System.Windows.Forms.TextBox replica;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label time_point_label;
        private System.Windows.Forms.Label layer_label;
        private System.Windows.Forms.Label formulation_label;
        private System.Windows.Forms.Label compound_label;
        private System.Windows.Forms.Label replica_label;
        private System.Windows.Forms.Button button_start;
        private System.Windows.Forms.RadioButton study_type_2;
        private System.Windows.Forms.RadioButton study_type_1;
        private System.Windows.Forms.DataGridView table;
        private System.Windows.Forms.Button generate_table;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button_load;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox api_count_box;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn volume;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column6;
        private System.Windows.Forms.Button button3;
    }
}

