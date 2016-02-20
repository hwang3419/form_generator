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
        public Dictionary<string, string> sub_param_dict;
        public Dictionary<string, Dictionary<string, string>> param_dict;
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
            count = Int32.Parse(this.compound.Text.ToString());
            for (int i = 1; i <=count; i++) {
                this.table.Rows.Add("Compound"+i.ToString(),i, 1);
            }
            count = Int32.Parse(this.time_point.Text.ToString());
            for (int i = 1; i <= count; i++) 
                {
                    this.table.Rows.Add("Time Point" + i.ToString(), i , 1);
                }
            count = Int32.Parse(this.layer.Text.ToString());
            for (int i = 1; i <= count; i++)
                {
                    this.table.Rows.Add("Layer" + i.ToString(), i , 1);
                }
            count = Int32.Parse(this.formulation.Text.ToString());
            for (int i = 1; i <= count; i++)
                {
                    this.table.Rows.Add("Formulation" + i.ToString(), i , 1);
                }

            }

        private DataGridView create_new_table_template() {
            var new_table = new DataGridView();
            new_table.ColumnCount = 2;
            new_table.Columns[0].Name = "Internal Sample ID";
            new_table.Columns[1].Name = "External Sample ID";
            return new_table;
        }

        private Dictionary<string, Dictionary<string, string>> load_params() {
            var compound_dict = new Dictionary<string, string>();
            var layer_dict = new Dictionary<string, string>();
            var time_dict = new Dictionary<string, string>();
            var formulation_dict = new Dictionary<string, string>();
            var result_dict = new Dictionary<string, Dictionary<string, string>>();
            foreach (DataGridViewRow row in this.table.Rows)
            {
                if (row.Cells[0].Value == null)
                    { continue; }
                if (row.Cells[0].Value.ToString().StartsWith("Compound")) {
                    compound_dict[row.Cells[2].Value.ToString()] = row.Cells[1].Value.ToString();
                }
                else if (row.Cells[0].Value.ToString().StartsWith("Layer"))
                {
                    layer_dict[row.Cells[2].Value.ToString()] = row.Cells[1].Value.ToString();
                }
                else if (row.Cells[0].Value.ToString().StartsWith("Time"))
                {
                    time_dict[row.Cells[2].Value.ToString()] = row.Cells[1].Value.ToString();
                }
                else if (row.Cells[0].Value.ToString().StartsWith("Formulation"))
                {
                    formulation_dict[row.Cells[2].Value.ToString()] = row.Cells[1].Value.ToString();
                }
            }
            result_dict["compound"] = compound_dict;
            result_dict["time"] = time_dict;
            result_dict["formulation"] = formulation_dict;
            result_dict["layer"] = layer_dict;
            return result_dict;
        }

        private void generate_table_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < tabControl1.TabPages.Count; i++) {
                tabControl1.TabPages.RemoveAt(i);
            }
            var result_dict = load_params();
            Console.WriteLine(result_dict);
            int label_name = 1;
            var test = create_new_table_template();
            var tab_page = new TabPage();
            tab_page.Text = label_name.ToString();
            tab_page.Controls.Add(test);
            this.tabControl1.Controls.Add(tab_page);
            tabControl1.Refresh();
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
