using System;
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

namespace WindowsFormsApplication1
{
    public partial class Excel_Gen : Form
    {
        public int study_type;
        public Dictionary<string, string> sub_param_dict;
        public Dictionary<string, Dictionary<string, string>> param_dict;
        public Dictionary<string, DataGridView> output_dict;
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
                this.table.Rows.Add("Compound" + i.ToString(), i, i);
            }
            count = Int32.Parse(this.time_point.Text.ToString());
            for (int i = 1; i <= count; i++)
            {
                this.table.Rows.Add("Time Point" + i.ToString(), i, i);
            }
            count = Int32.Parse(this.layer.Text.ToString());
            for (int i = 1; i <= count; i++)
            {
                this.table.Rows.Add("Layer" + i.ToString(), i, "L");
            }
            count = Int32.Parse(this.formulation.Text.ToString());
            for (int i = 1; i <= count; i++)
            {
                this.table.Rows.Add("Formulation" + i.ToString(), i, i);
            }

        }

        private DataGridView create_new_table_template(Dictionary<string, string> compound_dict)
        {
            var new_table = new DataGridView();
            new_table.ColumnCount = 2 + (int) compound_dict.LongCount();
            new_table.Columns[0].Name = "Internal Sample ID";
            new_table.Columns[0].Width = 200;
            new_table.Columns[1].Name = "External Sample ID";
            new_table.Columns[1].Width = 200;
            int column_index = 2;
            foreach(KeyValuePair<string,string> kv in compound_dict)
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
                exlabel = ex_prefix + layer_entry.Value;
                ex_start = ex_factor * replica_int * time_int;
                for (int i = 1; i <= time_int; i++)
                {
                    for (int j = 1; j <= replica_int; j++)
                    {
                        ex_start += 1;
                        ex_count = ex_start;
                        local_table.Rows.Add(inlabel + "-" + i.ToString() + "-" + j, exlabel + ex_count.ToString());
                    }

                }
                ex_start = ex_start + replica_int * time_int * formulation_int;
            }

            output_dict[local_tabpage.Text] = local_table;
            local_tabpage.Controls.Add(local_table);
            local_table.AutoResizeRows();
            return local_tabpage;
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
                
                var collection = new Microsoft.Office.Interop.Excel.Worksheet[output_dict.LongCount()+2];

                foreach (KeyValuePair<string,DataGridView> entry in output_dict)
                {
                    copyAlltoClipboard(entry.Value);



                    // xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(count);
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
                   //break;

                }
                

                

                // Save the excel file under the captured location from the SaveFileDialog
                xlWorkBook.SaveAs(sfd.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlexcel.DisplayAlerts = true;
                xlWorkBook.Close(true, misValue, misValue);
                xlexcel.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlexcel);

                // Clear Clipboard and DataGridView selection
                Clipboard.Clear();
                //dgvItems.ClearSelection();

                // Open the newly saved excel file
                if (File.Exists(sfd.FileName))
                    System.Diagnostics.Process.Start(sfd.FileName);
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
                ReadXls(filename, 1);
            }
        }

        public static List<List<string>> ReadXls(string filename, int index)
        {
           
            Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();
            object Missing = System.Reflection.Missing.Value;
            Excel.Workbook book = xls.Workbooks.Open(filename, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing);

            Excel.Worksheet sheet;//
            xls.Visible = false;//
            xls.DisplayAlerts = false;//

            try
            {
                sheet = (Excel.Worksheet)book.Worksheets.get_Item(index);
            }
            catch (Exception ex)//
            {
                Console.WriteLine(ex.Message);
                return null;
            }
            Console.WriteLine(sheet.Name);
            int row = sheet.UsedRange.Rows.Count;
            int col = sheet.UsedRange.Columns.Count;
            Excel.Range c1 = sheet.Cells[1, 1];
            Excel.Range c2 = sheet.Cells[row, col];
            var v = (Excel.Range)sheet.get_Range(c1,c2);
            var value = v.Value2;
            int count = 0;
            List<string> new_row = new List<string>();
            List<List<string>> result = new List<List<string>>();
            foreach (var item in value)
            {
                new_row.Add(item.ToString());
                count += 1;
                if (count% col == 0)
                {
                    result.Add(new_row);
                    new_row = new List<string>();
                }
            }
            book.Save();//
            book.Close(false, Missing, Missing);//
            xls.Quit();//

            sheet = null;
            book = null;
            xls = null;
            GC.Collect();
            return result;
        }


    }
}
