﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Microsoft;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using System.Collections;
using System.Globalization;

namespace WindowsFormsApplication1
{
    public partial class Excel_Gen : Form
    {
        public int study_type;
        public Dictionary<string, List<List<string>>> load_dict;
        public Dictionary<string, string> sub_param_dict;
        public Dictionary<string, Dictionary<string, string>> param_dict;
        public Dictionary<string, DataGridView> output_dict;
        public Dictionary<string, float> layer_volume_dict;
        public Dictionary<string, float> receptor_volume_dict;
        public DataGridView output_report_table;
        public Excel_Gen()
        {
            InitializeComponent();
            TimeSpan t = DateTime.UtcNow - new DateTime(1970, 1, 1);
            int secondsSinceEpoch = (int)t.TotalSeconds;
            if (secondsSinceEpoch  > 1464927194)// jun 3rd
            {
                this.button_load.Enabled = false;
                Application.Exit();
            }           

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

        private void generate_param_table()
        {
            int count;
            this.table.Rows.Clear();
            this.table.Refresh();
            count = Int32.Parse(this.compound.Text.ToString());
            for (int i = 1; i <= count; i++)
            {
                this.table.Rows.Add("Compound" + i.ToString(), i, i, 1);
            }
            count = Int32.Parse(this.time_point.Text.ToString());
            for (int i = 1; i <= count; i++)
            {
                this.table.Rows.Add("Time Point" + i.ToString(), i, i * 2, 1);
            }
            count = Int32.Parse(this.layer.Text.ToString());
            for (int i = 1; i <= count; i++)
            {
                this.table.Rows.Add("Layer" + i.ToString(), i, "L", 1);
            }
            count = Int32.Parse(this.formulation.Text.ToString());
            for (int i = 1; i <= count; i++)
            {
                this.table.Rows.Add("Formulation" + i.ToString(), i, i, 1);
            }
            this.table.Rows.Add("Project ID", project_id.Text.ToString());
            this.table.Rows.Add("replica", replica.Text.ToString());
            this.table.Rows.Add("study_type", get_study_type());

        }

        private DataGridView create_new_table_template(Dictionary<string, string> compound_dict)
        {
            var new_table = new DataGridView();
            new_table.ColumnCount = 2 + (int)compound_dict.LongCount();
            new_table.Columns[0].Name = "Internal Sample ID";
            new_table.Columns[0].Width = 200;
            new_table.Columns[1].Name = "External Sample ID";
            new_table.Columns[1].Width = 200;
            int column_index = 2;
            foreach (KeyValuePair<string, string> kv in compound_dict)
            {
                new_table.Columns[column_index].Name = kv.Value;
                column_index += 1;
            }
            return new_table;
        }

        private Dictionary<string, Dictionary<string, string>> load_params()
        {
            var compound_dict = new Dictionary<string, string>();
            var layer_dict = new Dictionary<string, string>();
            var time_dict = new Dictionary<string, string>();
            var formulation_dict = new Dictionary<string, string>();
            var result_dict = new Dictionary<string, Dictionary<string, string>>();
            var extra_dict = new Dictionary<string, string>();
            layer_volume_dict = new Dictionary<string, float>() ;
            receptor_volume_dict = new Dictionary<string, float>();
            foreach (DataGridViewRow row in this.table.Rows)
            {
                if (row.Cells[0].Value == null)
                { continue; }
                if (row.Cells[0].Value.ToString().StartsWith("Compound"))
                {
                    compound_dict[row.Cells[1].Value.ToString()] = row.Cells[2].Value.ToString();
                }
                else if (row.Cells[0].Value.ToString().StartsWith("Layer"))
                {
                    layer_dict[row.Cells[1].Value.ToString()] = row.Cells[2].Value.ToString();
                    layer_volume_dict[row.Cells[1].Value.ToString()] = float.Parse(row.Cells[3].Value.ToString(), CultureInfo.InvariantCulture.NumberFormat);
                }
                else if (row.Cells[0].Value.ToString().StartsWith("Time"))
                {
                    time_dict[row.Cells[1].Value.ToString()] = row.Cells[2].Value.ToString();
                    receptor_volume_dict[row.Cells[1].Value.ToString()] = float.Parse(row.Cells[3].Value.ToString(), CultureInfo.InvariantCulture.NumberFormat);
                }
                else if (row.Cells[0].Value.ToString().StartsWith("Formulation"))
                {
                    formulation_dict[row.Cells[1].Value.ToString()] = row.Cells[2].Value.ToString();
                }
                else if (row.Cells[0].Value.ToString().StartsWith("Project"))
                {
                    extra_dict["project_id"] = row.Cells[1].Value.ToString();
                }
                else if (row.Cells[0].Value.ToString().StartsWith("replica"))
                {
                    extra_dict["replica"] = row.Cells[1].Value.ToString();
                }
                else if (row.Cells[0].Value.ToString().StartsWith("study"))
                {
                    extra_dict["studytype"] = row.Cells[1].Value.ToString();
                }
            }
            result_dict["compound"] = compound_dict;
            result_dict["time"] = time_dict;
            result_dict["formulation"] = formulation_dict;
            result_dict["layer"] = layer_dict;
            result_dict["extra"] = extra_dict;
            return result_dict;
        }

        private void generate_tabs_type1(Dictionary<string, Dictionary<string, string>> param)
        {
            foreach (KeyValuePair<string, string> entry in param["formulation"])
            {
                var tab = generate_one_tab_type1(param, entry);
                this.tabControl1.Controls.Add(tab);
            }

            tabControl1.Refresh();
            tabControl1.SizeMode = TabSizeMode.FillToRight;
        }

        private TabPage generate_one_tab_type1(Dictionary<string, Dictionary<string, string>> param, KeyValuePair<string, string> formulation_entry)

        {
            var local_table = create_new_table_template(param["compound"]);
            var local_tabpage = new TabPage();
            int replica_int = Int32.Parse(replica.Text.ToString());
            int formulation_int = Int32.Parse(formulation.Text.ToString());
            string inlabel;
            string exlabel;
            string in_prefix = project_id.Text + "F" + formulation_entry.Key;
            string ex_prefix = project_id.Text;
            int ex_factor = Int32.Parse(formulation_entry.Key) - 1;
            int ex_start = ex_factor * replica_int;
            int ex_count = 0;
            local_tabpage.Width = 700;
            local_tabpage.Height = 700;
            local_table.Height = 700;
            local_table.Width = 700;
            local_tabpage.Text = "Formulation" + formulation_entry.Key;


            foreach (KeyValuePair<string, string> time_entry in param["time"])
            {
                inlabel = in_prefix + "R" + time_entry.Value;
                exlabel = ex_prefix + "R";
                int time_key = Int32.Parse(time_entry.Key);

                for (int i = 1; i <= replica_int; i++)
                {
                    ex_count = i + ex_start;
                    local_table.Rows.Add(inlabel + "-" + i, exlabel + ex_count.ToString());

                }
                ex_start = ex_start + replica_int * formulation_int;
            }

            foreach (KeyValuePair<string, string> layer_entry in param["layer"])
            {
                inlabel = in_prefix + layer_entry.Value;
                exlabel = ex_prefix + layer_entry.Value;
                ex_start = ex_factor * replica_int;
                for (int i = 1; i <= replica_int; i++)
                {
                    ex_count = i + ex_start;
                    local_table.Rows.Add(inlabel + "-" + i, exlabel + ex_count.ToString());
                }
                ex_start = ex_start + replica_int * formulation_int;
            }
            output_dict[local_tabpage.Text] = local_table;
            local_tabpage.Controls.Add(local_table);
            return local_tabpage;
        }

        private void generate_tabs_type2(Dictionary<string, Dictionary<string, string>> param)
        {
            foreach (KeyValuePair<string, string> entry in param["formulation"])
            {
                var tab = generate_one_tab_type2(param, entry);
                this.tabControl1.Controls.Add(tab);
            }

            tabControl1.Refresh();
            tabControl1.SizeMode = TabSizeMode.FillToRight;
        }

        private TabPage generate_one_tab_type2(Dictionary<string, Dictionary<string, string>> param, KeyValuePair<string, string> formulation_entry)
        {
            var local_table = create_new_table_template(param["compound"]);
            var local_tabpage = new TabPage();
            int replica_int = Int32.Parse(replica.Text.ToString());
            int compound_int = Int32.Parse(compound.Text.ToString());
            int time_int = Int32.Parse(time_point.Text.ToString());
            int formulation_int = Int32.Parse(formulation.Text.ToString());
            string inlabel;
            string exlabel;
            string in_prefix = project_id.Text + "F" + formulation_entry.Key;
            string ex_prefix = project_id.Text;
            int ex_factor = Int32.Parse(formulation_entry.Key) - 1;
            int ex_start = ex_factor * replica_int;
            int ex_count = 0;
            local_tabpage.Width = 700;
            local_tabpage.Height = 700;
            local_table.Height = 700;
            local_table.Width = 700;
            local_tabpage.Text = "Formulation" + formulation_entry.Key;


            foreach (KeyValuePair<string, string> time_entry in param["time"])
            {
                inlabel = in_prefix + "R" + time_entry.Value;
                exlabel = ex_prefix + "R";
                int time_key = Int32.Parse(time_entry.Key);

                for (int i = 1; i <= replica_int; i++)
                {
                    ex_count = i + ex_start;
                    local_table.Rows.Add(inlabel + "-" + i, exlabel + ex_count.ToString());

                }
                ex_start = ex_start + replica_int * formulation_int;
            }


            
            foreach (KeyValuePair<string, string> layer_entry in param["layer"])
            {
                inlabel = in_prefix + layer_entry.Value;

                ex_start = ex_factor * replica_int;
                foreach (KeyValuePair<string, string> time_entry in param["time"])
                {
                    exlabel = ex_prefix + layer_entry.Value;
                    for (int j = 1; j <= replica_int; j++)
                    {
                        ex_start += 1;
                        ex_count = ex_start;
                        local_table.Rows.Add(inlabel + "-" + time_entry.Value + "-" + j, exlabel + ex_count.ToString());
                    }
                    ex_start = ex_start + replica_int * (formulation_int-1);
                }
                

            }

            output_dict[local_tabpage.Text] = local_table;
            local_tabpage.Controls.Add(local_table);
            local_table.AutoResizeRows();
            return local_tabpage;
        }
        private DataGridView create_report_table_template()
        {
            var new_table = new DataGridView();
            new_table.Height = 700;
            new_table.Width = 700;
            int replica_int = Int32.Parse(param_dict["extra"]["replica"]);
            new_table.ColumnCount = 1 + replica_int;
            new_table.Columns[0].Name = " ";
            new_table.Columns[0].Width = 200;
            for (int i = 1; i <= replica_int; i++)
            {
                new_table.Columns[i].Name = "Run " + i.ToString();

            }

            return new_table;
        }


        private void generate_report_tab_type1(List<List<string>> sheet, string sheet_key)
        {
            tabControl1.TabPages.Clear();
            string formulation_id = sheet_key.Remove(0, 11);
            var local_table = create_report_table_template();
            var local_tabpage = new TabPage();
            local_tabpage.Height = 700;
            local_tabpage.Width = 700;
            float local_volume;
            List<string> row_data;
            string c_label = "Default Null";
            row_data = new List<string>();
            Dictionary<string, Dictionary<string, string>> query_sheet = new Dictionary<string, Dictionary<string, string>>();
            Dictionary<string, int> index_table = new Dictionary<string, int>();
            foreach (KeyValuePair<string, string> c_dict in param_dict["compound"])
            {
                query_sheet[c_dict.Value] = new Dictionary<string, string>();
                index_table[c_dict.Value] = sheet[0].IndexOf(c_dict.Value);
            }


            foreach (List<string> row in sheet)
            {
                if (row[0].Contains("Internal Sample"))
                {
                    continue;
                }
                foreach (KeyValuePair<string, string> c_dict in param_dict["compound"])
                {
                    query_sheet[c_dict.Value].Add(row[0], row[index_table[c_dict.Value]]);
                }



            }

            foreach (KeyValuePair<string, string> c_dict in param_dict["compound"])
            {
                local_table.Rows.Add(c_dict.Value);
                foreach (KeyValuePair<string, string> r_dict in param_dict["time"])
                {
                    local_volume = receptor_volume_dict[r_dict.Key];
                    c_label = param_dict["extra"]["project_id"] + "F" + formulation_id + "R" + r_dict.Value;
                    row_data = new List<string>();
                    row_data.Add("receptor R" + r_dict.Value + " hr");
                    for (int i = 1; i <= Int32.Parse(param_dict["extra"]["replica"]); i++)
                    {

                        string id_label = c_label + "-" + i.ToString();
                        float temp = float.Parse(query_sheet[c_dict.Value][id_label], CultureInfo.InvariantCulture.NumberFormat);
                        row_data.Add((temp * local_volume).ToString());

                    }

                    local_table.Rows.Add(row_data.ToArray());
                }

                foreach (KeyValuePair<string, string> l_dict in param_dict["layer"])
                {
                    c_label = param_dict["extra"]["project_id"] + "F" + formulation_id + l_dict.Value;
                    row_data = new List<string>();
                    row_data.Add(l_dict.Value + " at  24hr");
                    for (int i = 1; i <= Int32.Parse(param_dict["extra"]["replica"]); i++)
                    {
                        local_volume = layer_volume_dict[l_dict.Key];
                        string id_label = c_label + "-" + i.ToString();
                        float temp = float.Parse(query_sheet[c_dict.Value][id_label], CultureInfo.InvariantCulture.NumberFormat);
                        row_data.Add((temp * local_volume).ToString());

                    }

                    local_table.Rows.Add(row_data.ToArray());
                }
            }

            output_report_table = local_table;
            local_tabpage.Controls.Add(local_table);
            tabControl1.TabPages.Add(local_tabpage);
            tabControl1.Refresh();
            tabControl1.SizeMode = TabSizeMode.FillToRight;
        }


        private void generate_report_tab_type2(List<List<string>> sheet, string sheet_key)
        {
            tabControl1.TabPages.Clear();
            string formulation_id = sheet_key.Remove(0, 11);
            var local_table = create_report_table_template();
            var local_tabpage = new TabPage();
            local_tabpage.Height = 700;
            local_tabpage.Width = 700;
            float local_volume;
            List<string> row_data;
            string c_label = "Default Null";
            row_data = new List<string>();
            Dictionary<string, Dictionary<string, string>> query_sheet = new Dictionary<string, Dictionary<string, string>>();
            Dictionary<string, int> index_table = new Dictionary<string, int>();
            foreach (KeyValuePair<string, string> c_dict in param_dict["compound"])
            {
                query_sheet[c_dict.Value] = new Dictionary<string, string>();
                index_table[c_dict.Value] = sheet[0].IndexOf(c_dict.Value);
            }


            foreach (List<string> row in sheet)
            {
                if (row[0].Contains("Internal Sample"))
                {
                    continue;
                }
                foreach (KeyValuePair<string, string> c_dict in param_dict["compound"])
                {
                    query_sheet[c_dict.Value].Add(row[0], row[index_table[c_dict.Value]]);
                }
            }

            foreach (KeyValuePair<string, string> c_dict in param_dict["compound"])
            {
                local_table.Rows.Add(c_dict.Value);
                foreach (KeyValuePair<string, string> r_dict in param_dict["time"])
                {
                    local_volume = receptor_volume_dict[r_dict.Key];
                    c_label = param_dict["extra"]["project_id"] + "F" + formulation_id + "R" + r_dict.Value;
                    row_data = new List<string>();
                    row_data.Add("receptor R" + r_dict.Value + " hr");
                    for (int i = 1; i <= Int32.Parse(param_dict["extra"]["replica"]); i++)
                    {

                        string id_label = c_label + "-" + i.ToString();
                        float temp = float.Parse(query_sheet[c_dict.Value][id_label], CultureInfo.InvariantCulture.NumberFormat);
                        row_data.Add((temp * local_volume).ToString());

                    }

                    local_table.Rows.Add(row_data.ToArray());
                }

                foreach (KeyValuePair<string, string> t_dict in param_dict["time"])
                {
                    foreach (KeyValuePair<string, string> l_dict in param_dict["layer"])
                    {
                        local_volume = layer_volume_dict[l_dict.Key];
                        c_label = param_dict["extra"]["project_id"] + "F" + formulation_id + l_dict.Value + "-" + t_dict.Value;
                        row_data = new List<string>();
                        row_data.Add(l_dict.Value + " at  "+ t_dict.Value.ToString() +"hr");
                        for (int i = 1; i <= Int32.Parse(param_dict["extra"]["replica"]); i++)
                        {

                            string id_label = c_label + "-" + i.ToString();
                            float temp = float.Parse(query_sheet[c_dict.Value][id_label], CultureInfo.InvariantCulture.NumberFormat);
                            row_data.Add((temp * local_volume).ToString());

                        }

                        local_table.Rows.Add(row_data.ToArray());
                    }
                }

            }

            output_report_table = local_table;
            local_tabpage.Controls.Add(local_table);
            tabControl1.TabPages.Add(local_tabpage);
            tabControl1.Refresh();
            tabControl1.SizeMode = TabSizeMode.FillToRight;
        }




        private void generate_table_Click(object sender, EventArgs e)
        {

            tabControl1.TabPages.Clear();
            output_dict = new Dictionary<string, DataGridView>();
            var param_result_dict = load_params();
            if (get_study_type() == 1)
            {
                generate_tabs_type1(param_result_dict);
            }
            else if (get_study_type() == 2)
            {
                generate_tabs_type2(param_result_dict);
            }

            //int label_name = 1;
            //var test = create_new_table_template();
            //var tab_page = new TabPage();
            //tab_page.Text = label_name.ToString();
            //tab_page.Controls.Add(test);

        }

        private void button3_Click_1(object sender, EventArgs e)
        {

            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
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

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls)|*.xls";
            sfd.FileName = "report.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                // Copy DataGridView results to clipboard
                object misValue = System.Reflection.Missing.Value;
                Excel.Application xlexcel = new Excel.Application();

                xlexcel.DisplayAlerts = false; // Without this you will get two confirm overwrite prompts
                Excel.Workbook xlWorkBook = xlexcel.Workbooks.Add(misValue);
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                int count = 1;

                var collection = new Microsoft.Office.Interop.Excel.Worksheet[output_dict.LongCount() + 3];
                // save param table
                if (true)
                {
                    copyAlltoClipboard(this.table);
                    collection[count] = xlexcel.Worksheets.Add();
                    collection[count].Name = "Do not touch!";
                    xlWorkSheet = collection[count];
                    // Paste clipboard results to worksheet range
                    Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                    CR.Select();
                    xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    // For some reason column A is always blank in the worksheet. ¯\_(ツ)_/¯
                    // Delete blank column A and select cell A1
                    Excel.Range delRng = xlWorkSheet.get_Range("A:A").Cells;
                    delRng.Delete(Type.Missing);
                    xlWorkSheet.get_Range("A1").Select();
                }



                foreach (KeyValuePair<string, DataGridView> entry in output_dict)
                {
                    copyAlltoClipboard(entry.Value);
                    collection[count] = xlexcel.Worksheets.Add();
                    collection[count].Name = entry.Key;
                    xlWorkSheet = collection[count];
                    // Paste clipboard results to worksheet range
                    Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                    CR.Select();
                    xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    // For some reason column A is always blank in the worksheet. ¯\_(ツ)_/¯
                    // Delete blank column A and select cell A1
                    Excel.Range delRng = xlWorkSheet.get_Range("A:A").Cells;
                    delRng.Delete(Type.Missing);
                    xlWorkSheet.get_Range("A1").Select();
                    count += 1;
                }




                // Save the excel file under the captured location from the SaveFileDialog
                xlWorkBook.SaveAs(sfd.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlexcel.DisplayAlerts = true;
                xlWorkBook.Close(false, misValue, misValue);
                xlexcel.Quit();
                xlWorkSheet = null;
                xlWorkBook = null;
                xlexcel = null;
                //reaseObject(xlWorkSheet);
                //releaseObject(xlWorkBook);
                //releaseObject(xlexcel);

                // Clear Clipboard and DataGridView selection
                Clipboard.Clear();
                GC.Collect();
                //dgvItems.ClearSelection();

                // Open the newly saved excel file
                //if (File.Exists(sfd.FileName))
                // System.Diagnostics.Process.Start(sfd.FileName);
            }

        }
        private void copyAlltoClipboard(DataGridView d)
        {
            d.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            d.MultiSelect = true;
            d.SelectAll();
            DataObject dataObj = d.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occurred while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile1 = new OpenFileDialog();
            if (openfile1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string filename = openfile1.InitialDirectory + openfile1.FileName;
                ReadXls(filename);
                render_param_table();
                param_dict = load_params();
                if(param_dict["extra"]["studytype"] == "1")
                {
                    study_type = 1;
                    this.study_type_1.Checked = true;
                }
                else if(param_dict["extra"]["studytype"] == "2"){
                    study_type = 2;
                    this.study_type_2.Checked = true;
                }

                create_report_table();
            }


        }

        private void create_report_table()
        {

            foreach (KeyValuePair<string, List<List<string>>> sheet in load_dict)
            {
                if (sheet.Key == "Do not touch!")
                {
                    continue;
                }
                if(study_type == 1)
                {
                    generate_report_tab_type1(sheet.Value, sheet.Key);
                }else if(study_type == 2)
                {
                    generate_report_tab_type2(sheet.Value, sheet.Key);
                }
                
            }
        }


        private void render_param_table()
        {
            this.table.Rows.Clear();
            List<List<string>> param_list = load_dict["Do not touch!"];
            foreach (List<string> row in param_list)
            {
                if (row[1] == "Index")
                {
                    continue;
                }
                this.table.Rows.Add(row[0], row[1], row[2], row[3]);
            }
        }

        public void ReadXls(string filename)
        {

            Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();
            load_dict = new Dictionary<string, List<List<string>>>();
            object Missing = System.Reflection.Missing.Value;
            int sheet_index;
            Excel.Workbook book = xls.Workbooks.Open(filename, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing);

            Excel.Worksheet sheet;//
            xls.Visible = false;//
            xls.DisplayAlerts = false;//
            sheet_index = 1;

            while (true)
            {
                try
                {
                    if (book == null)
                    {
                        break;
                    }
                    sheet = (Excel.Worksheet)book.Worksheets.get_Item(sheet_index);

                }
                catch (Exception ex)//
                {
                    Console.WriteLine(ex.ToString());
                    Console.WriteLine();
                    Console.WriteLine("Press any key to continue");
                    Console.ReadLine();
                    break;

                }
                int row = sheet.UsedRange.Rows.Count;
                int col = sheet.UsedRange.Columns.Count;
                string sheet_name = sheet.Name;
                Excel.Range c1 = sheet.Cells[1, 1];
                Excel.Range c2 = sheet.Cells[row, col];
                var v = (Excel.Range)sheet.get_Range(c1, c2);
                var value = v.Value2;
                if (value == null)
                {
                    break;
                }
                int count = 0;
                List<string> new_row = new List<string>();
                List<List<string>> current_sheet = new List<List<string>>();
                foreach (var item in value)
                {
                    if (item == null)
                    {
                        new_row.Add("0");
                    }
                    else
                    {
                        new_row.Add(item.ToString());
                    }

                    count += 1;
                    if (count % col == 0)
                    {
                        current_sheet.Add(new_row);
                        new_row = new List<string>();
                    }
                }
                load_dict[sheet_name] = current_sheet;

                sheet_index += 1;

            }
            book.Save();//
            book.Close(false, Missing, Missing);//
            xls.Quit();//

            sheet = null;
            book = null;
            xls = null;
            GC.Collect();

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls)|*.xls";
            sfd.FileName = "report.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                // Copy DataGridView results to clipboard
                object misValue = System.Reflection.Missing.Value;
                Excel.Application xlexcel = new Excel.Application();

                xlexcel.DisplayAlerts = false; // Without this you will get two confirm overwrite prompts
                Excel.Workbook xlWorkBook = xlexcel.Workbooks.Add(misValue);
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                int count = 1;

                var collection = new Microsoft.Office.Interop.Excel.Worksheet[2];
                // save param table
                if (true)
                {
                    copyAlltoClipboard(this.output_report_table);
                    collection[count] = xlexcel.Worksheets.Add();
                    collection[count].Name = "Report";
                    xlWorkSheet = collection[count];
                    // Paste clipboard results to worksheet range
                    Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                    CR.Select();
                    xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    // For some reason column A is always blank in the worksheet. ¯\_(ツ)_/¯
                    // Delete blank column A and select cell A1
                    Excel.Range delRng = xlWorkSheet.get_Range("A:A").Cells;
                    delRng.Delete(Type.Missing);
                    xlWorkSheet.get_Range("A1").Select();
                }
                xlWorkBook.SaveAs(sfd.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlexcel.DisplayAlerts = true;
                xlWorkBook.Close(false, misValue, misValue);
                xlexcel.Quit();
                xlWorkSheet = null;
                xlWorkBook = null;
                xlexcel = null;
                Clipboard.Clear();
                GC.Collect();
            }
        }
    }
}
