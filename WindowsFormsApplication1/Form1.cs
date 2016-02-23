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
           
            Console.Write(study_type);
            Console.Write(this.layer.Text.ToString());
            generate_param_table();
        }

        private int get_study_type()
        {
            bool isChecked = this.study_type_1.Checked;
            if (isChecked)
                study_type = 1;
            else
                study_type = 2;
            return study_type;
        }

        private void generate_param_table() {
            int count;
            this.table.Rows.Clear();
            this.table.Refresh();
            count = Int32.Parse(this.compound.Text.ToString());
            for (int i = 1; i <=count; i++) {
                this.table.Rows.Add("Compound"+i.ToString(),i, i);
            }
            count = Int32.Parse(this.time_point.Text.ToString());
            for (int i = 1; i <= count; i++) 
                {
                    this.table.Rows.Add("Time Point" + i.ToString(), i , i);
                }
            count = Int32.Parse(this.layer.Text.ToString());
            for (int i = 1; i <= count; i++)
                {
                    this.table.Rows.Add("Layer" + i.ToString(), i , "L");
                }
            count = Int32.Parse(this.formulation.Text.ToString());
            for (int i = 1; i <= count; i++)
                {
                    this.table.Rows.Add("Formulation" + i.ToString(), i , i);
                }

            }

        private DataGridView create_new_table_template() {
            var new_table = new DataGridView();
            new_table.ColumnCount = 2;
            new_table.Columns[0].Name = "Internal Sample ID";
            new_table.Columns[0].Width = 200;
            new_table.Columns[1].Name = "External Sample ID";
            new_table.Columns[1].Width = 200;
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
                    compound_dict[row.Cells[1].Value.ToString()] = row.Cells[2].Value.ToString();
                }
                else if (row.Cells[0].Value.ToString().StartsWith("Layer"))
                {
                    layer_dict[row.Cells[1].Value.ToString()] = row.Cells[2].Value.ToString();
                }
                else if (row.Cells[0].Value.ToString().StartsWith("Time"))
                {
                    time_dict[row.Cells[1].Value.ToString()] = row.Cells[2].Value.ToString();
                }
                else if (row.Cells[0].Value.ToString().StartsWith("Formulation"))
                {
                    formulation_dict[row.Cells[1].Value.ToString()] = row.Cells[2].Value.ToString();
                }
            }
            result_dict["compound"] = compound_dict;
            result_dict["time"] = time_dict;
            result_dict["formulation"] = formulation_dict;
            result_dict["layer"] = layer_dict;
            return result_dict;
        }

        private void generate_tabs_type1(Dictionary<string, Dictionary<string, string>>  param) {
            foreach (KeyValuePair<string, string> entry in param["formulation"]) {
                var tab = generate_one_tab_type1(param, entry);
                this.tabControl1.Controls.Add(tab);
            }
           
            tabControl1.Refresh();
            tabControl1.SizeMode = TabSizeMode.FillToRight;
        }

        private TabPage generate_one_tab_type1(Dictionary<string, Dictionary<string, string>> param, KeyValuePair<string, string> formulation_entry)

        {
            var local_table = create_new_table_template();
            var local_tabpage = new TabPage(); 
            int replica_int = Int32.Parse(replica.Text.ToString());
            int formulation_int = Int32.Parse(formulation.Text.ToString());
            string inlabel;
            string exlabel;
            string in_prefix = project_id.Text + "F"+ formulation_entry.Key;
            string ex_prefix = project_id.Text;
            int ex_factor = Int32.Parse(formulation_entry.Key) - 1;
            int ex_start = ex_factor * replica_int;
            int ex_count =0;
            local_tabpage.Width = 800;
            local_table.Width = 1000;
            local_tabpage.Text = "Formulation" + formulation_entry.Key;

           
            foreach (KeyValuePair<string, string> time_entry in param["time"])
            {
                inlabel = in_prefix +"R"+time_entry.Value;
                exlabel = ex_prefix + "R";
                int time_key = Int32.Parse(time_entry.Key);
                
                for (int i=1;i<= replica_int;i++)
                {
                    ex_count =  i+ ex_start;
                    local_table.Rows.Add(inlabel + "-" + i, exlabel + ex_count.ToString());
                    
                }
                ex_start = ex_start + replica_int*formulation_int;
            }
            ex_start = ex_factor * replica_int;
            foreach (KeyValuePair<string, string> layer_entry in param["layer"])
            {
                inlabel = in_prefix + layer_entry.Value;
                exlabel = ex_prefix + layer_entry.Value;
                for (int i = 1; i <= replica_int; i++)
                {
                    ex_count = i + ex_start;
                    local_table.Rows.Add(inlabel + "-" + i, exlabel + ex_count.ToString());
                }
                ex_start = ex_start + replica_int * formulation_int;
            }
                local_tabpage.Controls.Add(local_table);
            return local_tabpage;
        }

        private void generate_table_Click(object sender, EventArgs e)
        {

            tabControl1.TabPages.Clear();
            var param_result_dict = load_params();
            if (get_study_type() == 1)
            {
                generate_tabs_type1(param_result_dict);
            }
            else if (get_study_type() == 2) {
            }
            
            int label_name = 1;
            var test = create_new_table_template();
            var tab_page = new TabPage();
            tab_page.Text = label_name.ToString();
            tab_page.Controls.Add(test);
            
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
