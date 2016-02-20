using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Excel_Gen : Form
    {
        public int study_type;
        public Excel_Gen()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

            
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            bool isChecked = this.study_type_1.Checked;
            if (isChecked)
                study_type = 1;
            else
                study_type = 2;
            Console.Write(study_type);
            Console.Write(this.layer.Text.ToString());
            generate_param_table();
        }

        private void generate_param_table() {
            int count;
            this.table.Rows.Clear();
            this.table.Refresh();
            this.table.Rows.Add("Compound Number", "Compound Name");
            count = Int32.Parse(this.compound.Text.ToString());
            for (var i = 0; i < count; i++) {
                this.table.Rows.Add(i+1, "");
            }
            this.table.Rows.Add("Time Point Number", "Time Point");
            count = Int32.Parse(this.time_point.Text.ToString());
            for (var i = 0; i < count; i++)
            {
                this.table.Rows.Add(i + 1, "");
            }
            this.table.Rows.Add("Layer Number", "Layer Name");
            count = Int32.Parse(this.layer.Text.ToString());
            for (var i = 0; i < count; i++)
            {
                this.table.Rows.Add(i + 1, "");
            }
            this.table.Rows.Add("Formulation Number", "Formulation Name");
            count = Int32.Parse(this.formulation.Text.ToString());
            for (var i = 0; i < count; i++)
            {
                this.table.Rows.Add(i + 1, "");
            }

        }

        private void Excel_Gen_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void table_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void replica_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
