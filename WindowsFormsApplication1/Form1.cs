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
        public Dictionary<string, DataGridView> label2_output_dict;
        public Dictionary<string, float> layer_volume_dict;
        public Dictionary<string, float> receptor_volume_dict;
        public DataGridView output_report_table;
        public Dictionary<string, List<string>> formulation_extra_dict;
        public Dictionary<string, float> mass_balance_dict;
        public Dictionary<string, List<string>> label2_dict;
        public Excel_Gen()
        {
            InitializeComponent();
            TimeSpan t = DateTime.UtcNow - new DateTime(1970, 1, 1);
            int secondsSinceEpoch = (int)t.TotalSeconds;
            if (secondsSinceEpoch > 1464927194)// jun 3rd
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
                this.table.Rows.Add("Compound" + i.ToString(), i, i, 1, 0, 0);
            }
            count = Int32.Parse(this.time_point.Text.ToString());
            for (int i = 1; i <= count; i++)
            {
                this.table.Rows.Add("Time Point" + i.ToString(), i, i * 2, 1, 0, 0);
            }
            count = Int32.Parse(this.layer.Text.ToString());
            for (int i = 1; i <= count; i++)
            {
                this.table.Rows.Add("Layer" + i.ToString(), i, "L", 1, 0, 0);
            }
            count = Int32.Parse(this.formulation.Text.ToString());
            for (int i = 1; i <= count; i++)
            {
                this.table.Rows.Add("Formulation" + i.ToString(), i, i, 1, 0, 0);
            }
            count = Int32.Parse(this.api_count_box.Text.ToString());
            for (int i = 1; i <= count; i++)
            {
                this.table.Rows.Add("API_" + i.ToString(), i, i, 1, 0, 0);
            }
            this.table.Rows.Add("Project ID", project_id.Text.ToString());
            this.table.Rows.Add("replica", replica.Text.ToString());
            this.table.Rows.Add("study_type", get_study_type());
            this.table.Rows.Add("api", api_count_box.Text.ToString());

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
            var api_con_dict = new Dictionary<string, string>();
            var result_dict = new Dictionary<string, Dictionary<string, string>>();
            var extra_dict = new Dictionary<string, string>();
            List<string> formulation_extra_list = new List<string>();
            layer_volume_dict = new Dictionary<string, float>();
            receptor_volume_dict = new Dictionary<string, float>();
            formulation_extra_dict = new Dictionary<string, List<string>>();
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
                    formulation_extra_dict[row.Cells[1].Value.ToString()] = new List<string> { row.Cells[4].Value.ToString(), row.Cells[5].Value.ToString() };
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
                else if (row.Cells[0].Value.ToString().StartsWith("api"))
                {
                    extra_dict["api"] = row.Cells[1].Value.ToString();
                }
                else if (row.Cells[0].Value.ToString().StartsWith("API_"))
                {
                    api_con_dict[row.Cells[1].Value.ToString()] = row.Cells[2].Value.ToString();
                }
            }
            result_dict["compound"] = compound_dict;
            result_dict["time"] = time_dict;
            result_dict["formulation"] = formulation_dict;
            result_dict["layer"] = layer_dict;
            result_dict["api_con"] = api_con_dict;
            result_dict["extra"] = extra_dict;
            return result_dict;
        }

        private void generate_tabs_type1(Dictionary<string, Dictionary<string, string>> param)
        {

            foreach (KeyValuePair<string, string> entry in param["formulation"])
            {
                var tab = generate_one_tab_type1(param, entry);
                //order_list.Add(tab);
                this.tabControl1.TabPages.Insert(tabControl1.TabPages.Count, tab);
            }
            //for(int i = 0; i < order_list.Count(); i++) { }

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
            List<string> internal_label = new List<string>();
            List<string> external_label = new List<string>();
            List<string> layer_label = new List<string>();

            foreach (KeyValuePair<string, string> time_entry in param["time"])
            {
                inlabel = in_prefix + "R" + time_entry.Value;
                exlabel = ex_prefix + "R";
                int time_key = Int32.Parse(time_entry.Key);

                for (int i = 1; i <= replica_int; i++)
                {
                    ex_count = i + ex_start;
                    local_table.Rows.Add(inlabel + "-" + i, exlabel + ex_count.ToString());
                    internal_label.Add(inlabel + "-" + i);
                    external_label.Add(exlabel + ex_count.ToString());
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
                    internal_label.Add(inlabel + "-" + i);
                    external_label.Add(exlabel + ex_count.ToString());
                    layer_label.Add(inlabel + "-" + i);
                }
                ex_start = ex_start + replica_int * formulation_int;
            }
            output_dict[local_tabpage.Text] = local_table;
            local_tabpage.Controls.Add(local_table);
            internal_label = internal_label.Concat(layer_label).Concat(external_label).ToList();
            label2_dict[local_tabpage.Text] = internal_label;
            return local_tabpage;
        }

        private void generate_tabs_type2(Dictionary<string, Dictionary<string, string>> param)
        {
            foreach (KeyValuePair<string, string> entry in param["formulation"])
            {
                var tab = generate_one_tab_type2(param, entry);
                this.tabControl1.TabPages.Insert(tabControl1.TabPages.Count, tab);
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
                    ex_start = ex_start + replica_int * (formulation_int - 1);
                }


            }

            output_dict[local_tabpage.Text] = local_table;
            local_tabpage.Controls.Add(local_table);
            local_table.AutoResizeRows();
            return local_tabpage;
        }
        private List<string> get_report_header()
        {
            List<string> header = new List<string>();
            header.Add("Collection Point");
            int replica_int = Int32.Parse(param_dict["extra"]["replica"]);
            for (int i = 1; i <= replica_int; i++)
            {
                header.Add("Run  #" + i.ToString() + " (ng)");

            }
            header.Add("");
            header.Add("Average");
            header.Add("STD");
            header.Add("STD/AVG");
            return header;
        }

        private DataGridView create_report_table_template()
        {
            var new_table = new DataGridView();
            new_table.Height = 700;
            new_table.Width = 700;
            int replica_int = Int32.Parse(param_dict["extra"]["replica"]);
            int api_int = Int32.Parse(param_dict["extra"]["api"]);
            new_table.ColumnCount = 1 + replica_int + 4 + api_int;
            new_table.Columns[0].Name = " ";
            new_table.Columns[0].Width = 200;
            new_table.Rows.Add("Project Name:");
            new_table.Rows.Add("Prepared By:");
            new_table.Rows.Add("Project No.");
            new_table.Rows.Add("For");
            new_table.Rows.Add("Date");
            new_table.Rows.Add("");
            new_table.Rows.Add("Tissue diffusion area (cm2):");
            List<string> api_list = new List<string>();
            mass_balance_dict = new Dictionary<string, float>();
            api_list.Add("API");
            foreach (KeyValuePair<string, string> kv in param_dict["compound"])
            {
                api_list.Add(kv.Value);
            }
            new_table.Rows.Add(api_list.ToArray());
            List<string> formualation_header = new List<string> { "", "" };
            for (int i = 1; i <= api_int; i++)
            {
                formualation_header.Add("API Concentration " + i.ToString());
            }
            formualation_header.Add(" Dosed Amount / mg");
            formualation_header.Add("Dosed Amount for Each Cell/ mg");
            for (int i = 1; i <= api_int; i++)
            {
                formualation_header.Add(" Applied Amount of API" + i.ToString() + " / ng");
            }
            new_table.Rows.Add(formualation_header.ToArray());
            int count = 0;
            foreach (KeyValuePair<string, List<string>> kv in formulation_extra_dict)
            {
                List<string> rowdata = new List<string>();
                if (count == 0)
                {
                    count += 1;
                    rowdata.Add("Formulation");
                }
                else
                {
                    rowdata.Add("");
                }


                float temp1, temp2, temp3;
                rowdata.Add(param_dict["formulation"][kv.Key.ToString()]);
                float dose_before, dose_after, api_con;
                float.TryParse(kv.Value[0].ToString(), out dose_before);
                float.TryParse(kv.Value[1].ToString(), out dose_after);
                //  if (kv.Value[0].ToString().Contains(","))
                //  {
                //string[] api_con_list = kv.Value[0].Split(',');
                //string api_con_result="";
                foreach (KeyValuePair<string, string> kv2 in param_dict["api_con"])
                {
                    rowdata.Add(kv2.Value);
                }
                temp1 = (dose_before - dose_after) * 1000;
                rowdata.Add(temp1.ToString("0.0"));
                temp2 = temp1 / replica_int;
                mass_balance_dict[kv.Key.ToString()] = temp2;
                rowdata.Add(temp2.ToString("0.0"));
                foreach (KeyValuePair<string, string> kv2 in param_dict["api_con"])
                {
                    float.TryParse(kv2.Value.ToString(), out api_con);
                    temp3 = api_con * temp2 * 1000000;
                    rowdata.Add(temp3.ToString("0.0"));
                }
                /*   }
                   else
                   {
                       float.TryParse(kv.Value[0].ToString(), out api_con);
                       rowdata.Add(api_con.ToString());
                       temp1 = (dose_before - dose_after) * 1000;
                       rowdata.Add(temp1.ToString());
                       temp2 = temp1 / replica_int;
                       mass_balance_dict[kv.Key.ToString()] = temp2;
                       rowdata.Add(temp2.ToString());
                       temp3 = api_con * temp2 * 1000000;
                       rowdata.Add(temp3.ToString());
                   }*/
                new_table.Rows.Add(rowdata.ToArray());

            }

            new_table.Rows.Add("Tissue No.");
            new_table.Rows.Add("Age/Race/Gender");
            new_table.Rows.Add("Thickness/mm");
            api_list = new List<string>();
            api_list.Add("Time point");
            foreach (KeyValuePair<string, string> kv in param_dict["time"])
            {
                api_list.Add(kv.Value.ToString());
            }
            new_table.Rows.Add(api_list.ToArray());
            new_table.Rows.Add("Replicate", replica_int.ToString(), "Time Points", param_dict["time"].Count().ToString());
            new_table.Rows.Add("Note:");
            new_table.Rows.Add(" ");


            new_table.Rows.Add(get_report_header().ToArray());
            new_table.Rows.Add("Skin Tissue Information");
            return new_table;
        }

        private float stofloat(string param)
        {

            return float.Parse(param, CultureInfo.InvariantCulture.NumberFormat);
        }

        private List<string> append_data(List<string> data)
        {
            float sum = 0;
            double stdsum = 0;
            double std;
            float avg;
            long length = data.LongCount() - 1;
            for (int i = 1; i < data.LongCount(); i++)
            {
                float result;
                bool isNumber = float.TryParse(data[i], out result);
                if (isNumber)
                {
                    sum += result;
                }
                else
                {
                    data.Add("NA");
                    data.Add("NA");
                    data.Add("NA");
                    data.Add("NA");
                    return data;
                }

            }
            avg = sum / (length);
            for (int i = 1; i < data.LongCount(); i++)
            {
                stdsum += Math.Pow(stofloat(data[i]) - avg, 2);
            }
            std = Math.Pow(stdsum / length, 0.5);
            data.Add("");
            data.Add(avg.ToString("0.0"));
            data.Add(std.ToString("0.0"));
            data.Add((std / avg).ToString("0.0"));
            return data;
        }
        private string get_formulation_name(string sheetid)
        {
            string formulation_id = sheetid.Remove(0, 11);
            return param_dict["formulation"][formulation_id];
        }

        private void generate_report_tab_type1(List<List<string>> sheet, string sheet_key)
        {

            string formulation_id = sheet_key.Remove(0, 11);
            string formulation_realname = get_formulation_name(sheet_key);
            DataGridView local_table;
            if (output_report_table.RowCount < 1)
            {
                local_table = create_report_table_template();
            }
            else
            {
                local_table = output_report_table;
            }

            float local_volume;
            string sum_average;
            float sum_average_float;
            sum_average = "";
            sum_average_float = 0;
            List<string> row_data;
            string c_label = "NA";
            row_data = new List<string>();
            Dictionary<string, Dictionary<string, string>> query_sheet = new Dictionary<string, Dictionary<string, string>>();
            query_sheet.Clear();
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
            List<string> last_row_data = new List<string>();
            foreach (KeyValuePair<string, string> c_dict in param_dict["compound"])
            {
                local_table.Rows.Add("API", c_dict.Value);
                local_table.Rows.Add(formulation_realname);
                last_row_data = new List<string>();
                foreach (KeyValuePair<string, string> r_dict in param_dict["time"])
                {
                    local_volume = receptor_volume_dict[r_dict.Key];
                    c_label = param_dict["extra"]["project_id"] + "F" + formulation_id + "R" + r_dict.Value;
                    row_data = new List<string>();
                    row_data.Add("receptor at " + r_dict.Value + " hr");
                    for (int i = 1; i <= Int32.Parse(param_dict["extra"]["replica"]); i++)
                    {

                        string id_label = c_label + "-" + i.ToString();
                        float float_result;
                        bool isNumber = float.TryParse(query_sheet[c_dict.Value][id_label], out float_result);
                        if (isNumber)
                        {
                            row_data.Add((float_result * local_volume).ToString());
                        }
                        else
                        {
                            row_data.Add("NA");
                        }
                        //float temp = float.Parse(query_sheet[c_dict.Value][id_label], CultureInfo.InvariantCulture.NumberFormat);
                        //row_data.Add((temp * local_volume).ToString());

                    }
                    if (last_row_data.LongCount() == 0)
                    {
                        last_row_data = row_data.ToList();
                    }
                    else
                    {
                        for (int i = 1; i < last_row_data.LongCount(); i++)
                        {
                            float float_result_1;
                            bool isNumber_1 = float.TryParse(row_data[i], out float_result_1);
                            float float_result_2;
                            bool isNumber_2 = float.TryParse(last_row_data[i], out float_result_2);
                            if (isNumber_1 && isNumber_2)
                            {
                                row_data[i] = (float_result_1 + float_result_2).ToString();
                            }
                            else
                            {
                                row_data[i] = "NA";
                            }

                        }
                        last_row_data = row_data.ToList();
                    }
                    row_data = append_data(row_data);
                    local_table.Rows.Add(row_data.ToArray());
                    sum_average = row_data[row_data.Count() - 3];
                    float.TryParse(sum_average, out sum_average_float);
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
                        float float_result;
                        bool isNumber = float.TryParse(query_sheet[c_dict.Value][id_label], out float_result);
                        if (isNumber)
                        {
                            row_data.Add((float_result * local_volume).ToString());
                        }
                        else
                        {
                            row_data.Add("NA");
                        }


                    }
                    row_data = append_data(row_data);
                    local_table.Rows.Add(row_data.ToArray());
                    float result_float;
                    sum_average = row_data[row_data.Count() - 3];
                    if (float.TryParse(sum_average, out result_float))
                    {
                        sum_average_float += result_float;
                    }
                    else
                    {
                        sum_average = "NA";
                    }
                }


                if (sum_average == "NA")
                {
                    local_table.Rows.Add("Sum", "NA");

                }
                else
                {
                    local_table.Rows.Add("sum", sum_average_float.ToString("0.0"));
                    local_table.Rows.Add("mass balance");//, (sum_average_float / mass_balance_dict[formulation_id]).ToString());
                }
                local_table.Rows.Add("");
            }


            output_report_table = local_table;

        }


        private void generate_report_tab_type2(List<List<string>> sheet, string sheet_key)
        {
            string formulation_id = sheet_key.Remove(0, 11);
            string formulation_realname = get_formulation_name(sheet_key);
            DataGridView local_table;
            if (output_report_table.RowCount < 1)
            {
                local_table = create_report_table_template();
            }
            else
            {
                local_table = output_report_table;
            }
            float local_volume;
            List<string> row_data;
            string c_label = "Default Null";
            row_data = new List<string>();
            Dictionary<string, Dictionary<string, string>> query_sheet = new Dictionary<string, Dictionary<string, string>>();
            query_sheet.Clear();
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
                local_table.Rows.Add("API", c_dict.Value);
                local_table.Rows.Add(formulation_realname);
                foreach (KeyValuePair<string, string> r_dict in param_dict["time"])
                {
                    local_volume = receptor_volume_dict[r_dict.Key];
                    c_label = param_dict["extra"]["project_id"] + "F" + formulation_id + "R" + r_dict.Value;
                    row_data = new List<string>();
                    row_data.Add("receptor at " + r_dict.Value + " hr");
                    for (int i = 1; i <= Int32.Parse(param_dict["extra"]["replica"]); i++)
                    {

                        string id_label = c_label + "-" + i.ToString();
                        float float_result;
                        bool isNumber = float.TryParse(query_sheet[c_dict.Value][id_label], out float_result);
                        if (isNumber)
                        {
                            row_data.Add((float_result * local_volume).ToString());
                        }
                        else
                        {
                            row_data.Add("NA");
                        }

                    }
                    row_data = append_data(row_data);
                    local_table.Rows.Add(row_data.ToArray());


                    foreach (KeyValuePair<string, string> l_dict in param_dict["layer"])
                    {
                        local_volume = layer_volume_dict[l_dict.Key];
                        c_label = param_dict["extra"]["project_id"] + "F" + formulation_id + l_dict.Value + "-" + r_dict.Value;
                        row_data = new List<string>();
                        row_data.Add(l_dict.Value + " at  " + r_dict.Value.ToString() + "hr");
                        for (int i = 1; i <= Int32.Parse(param_dict["extra"]["replica"]); i++)
                        {

                            string id_label = c_label + "-" + i.ToString();
                            float float_result;
                            bool isNumber = float.TryParse(query_sheet[c_dict.Value][id_label], out float_result);
                            if (isNumber)
                            {
                                row_data.Add((float_result * local_volume).ToString());
                            }
                            else
                            {
                                row_data.Add("NA");
                            }
                            // float temp = float.Parse(query_sheet[c_dict.Value][id_label], CultureInfo.InvariantCulture.NumberFormat);
                            // row_data.Add((temp * local_volume).ToString());

                        }
                        row_data = append_data(row_data);
                        local_table.Rows.Add(row_data.ToArray());
                    }
                    local_table.Rows.Add("");
                }

                //foreach (KeyValuePair<string, string> t_dict in param_dict["time"])
                //{

                //}
                local_table.Rows.Add("");
            }

            output_report_table = local_table;
        }




        private void generate_table_Click(object sender, EventArgs e)
        {

            tabControl1.TabPages.Clear();
            output_dict = new Dictionary<string, DataGridView>();
            label2_dict = new Dictionary<string, List<string>>();
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

                Dictionary<int, DataGridView> output_dict_int = new Dictionary<int, DataGridView>();

                foreach (KeyValuePair<string, DataGridView> entry in output_dict)
                {
                    if (entry.Key.StartsWith("Formulation"))
                    {
                        int index_key = 0;
                        Int32.TryParse(entry.Key.Substring(11), out index_key);
                        output_dict_int[index_key] = entry.Value;
                    }

                }


                foreach (KeyValuePair<int, DataGridView> entry in output_dict_int.OrderBy(key => key.Key).Reverse())
                {
                    copyAlltoClipboard(entry.Value);
                    collection[count] = xlexcel.Worksheets.Add();
                    collection[count].Name = "Formulation" + entry.Key.ToString();
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
                if (param_dict["extra"]["studytype"] == "1")
                {
                    study_type = 1;
                    this.study_type_1.Checked = true;
                }
                else if (param_dict["extra"]["studytype"] == "2")
                {
                    study_type = 2;
                    this.study_type_2.Checked = true;
                }

                create_report_table();
            }


        }

        private void create_report_table()
        {
            tabControl1.TabPages.Clear();
            output_report_table = new DataGridView();
            var local_tabpage = new TabPage();
            local_tabpage.Height = 700;
            local_tabpage.Width = 700;
            List<string> fname_list = new List<string>();
            for (int i = 1; i <= param_dict["formulation"].Count(); i++)
            {
                fname_list.Add("Formulation" + i.ToString());
            }
            //foreach (KeyValuePair<string, List<List<string>>> sheet in load_dict)
            foreach (string fname in fname_list)
            {
                if (study_type == 1)
                {
                    generate_report_tab_type1(load_dict[fname], fname);
                }
                else if (study_type == 2)
                {
                    generate_report_tab_type2(load_dict[fname], fname);
                }

            }
            add_total_mass();
            local_tabpage.Controls.Add(output_report_table);
            tabControl1.TabPages.Add(local_tabpage);
            tabControl1.Refresh();
            tabControl1.SizeMode = TabSizeMode.FillToRight;
        }


        private void add_total_mass()
        {
            //float total = mass_balance_dict.Values.Sum();
            output_report_table.Rows.Add("Total Mass Balance");//, total.ToString());

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
                this.table.Rows.Add(row[0], row[1], row[2], row[3], row[4], row[5]);
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
        private DataGridView create_new_table2_template()
        {
            var new_table = new DataGridView();
            new_table.ColumnCount = 1;
            new_table.Columns[0].Name = "Sample ID";
            new_table.Columns[0].Width = 200;
            return new_table;
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click_2(object sender, EventArgs e)
        {
            this.tabControl1.TabPages.Clear();
            label2_output_dict = new Dictionary<string, DataGridView>();

            foreach (KeyValuePair<string, List<string>> entry in label2_dict)
            {
                var local_table = create_new_table2_template();
                var local_tabpage = new TabPage();
                for (int i = 0; i < entry.Value.Count(); i++)
                {
                    local_table.Rows.Add(entry.Value[i]);
                }
                label2_output_dict[entry.Key] = local_table;
                local_tabpage.Controls.Add(local_table);
                this.tabControl1.TabPages.Insert(tabControl1.TabPages.Count, local_tabpage);
            }
            //for(int i = 0; i < order_list.Count(); i++) { }

            tabControl1.Refresh();
            tabControl1.SizeMode = TabSizeMode.FillToRight;


            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls)|*.xls";
            sfd.FileName = "label2.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                // Copy DataGridView results to clipboard
                object misValue = System.Reflection.Missing.Value;
                Excel.Application xlexcel = new Excel.Application();

                xlexcel.DisplayAlerts = false; // Without this you will get two confirm overwrite prompts
                Excel.Workbook xlWorkBook = xlexcel.Workbooks.Add(misValue);
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                int count = 1;

                var collection = new Microsoft.Office.Interop.Excel.Worksheet[label2_dict.LongCount()+2];
                // save param table

                Dictionary<int, DataGridView> output_dict_int = new Dictionary<int, DataGridView>();

                foreach (KeyValuePair<string, DataGridView> entry in label2_output_dict)
                {
                    if (entry.Key.StartsWith("Formulation"))
                    {
                        int index_key = 0;
                        Int32.TryParse(entry.Key.Substring(11), out index_key);
                        output_dict_int[index_key] = entry.Value;
                    }

                }


                foreach (KeyValuePair<int, DataGridView> entry in output_dict_int.OrderBy(key => key.Key).Reverse())
                {
                    copyAlltoClipboard(entry.Value);
                    collection[count] = xlexcel.Worksheets.Add();
                    collection[count].Name = "Formulation" + entry.Key.ToString();
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
            }
        }
    }
}
