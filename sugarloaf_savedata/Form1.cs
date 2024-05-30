using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Xls;

namespace sugarloaf_savedata {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e) {
            DirectoryInfo lish_excel = new DirectoryInfo(Application.ExecutablePath.Replace(".EXE", ".exe").Replace("sugarloaf_savedata.exe", ""));
            FileInfo[] Files = lish_excel.GetFiles("*.xlsx");
            foreach (FileInfo file in Files) {
                if (file.Name.Contains("~")) continue;
                comboBox1.Items.Add(file.Name.Replace(file.Extension, ""));
            }
            FileInfo[] Filess = lish_excel.GetFiles("*.ods");
            foreach (FileInfo file in Filess) {
                if (file.Name.Contains("~")) continue;
                comboBox1.Items.Add(file.Name.Replace(file.Extension, ""));
            }
        }

        Workbook workbook = new Workbook();
        Worksheet worksheet;
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            string[] s_head = File.ReadAllText(comboBox1.Text + ".txt").Replace("\r\n", "|").Split('|');
            string[] s_step = File.ReadAllText(comboBox1.Text + "_step.txt").Replace("\r\n", "|").Split('|');
            try {
                workbook.LoadFromFile(comboBox1.Text + comboBox2.Text);
            } catch {
                MessageBox.Show("_ปิด excel ก่อน");
                return;
            }
            worksheet = workbook.Worksheets["info"];
            int page_excel = 7;
            string sheet_test = "";
            sheet_test = worksheet.GetText(page_excel, 2);
            try { worksheet = workbook.Worksheets[sheet_test]; } catch { MessageBox.Show("_กรุณาปิด excel ก่อน"); return; }
            worksheet = workbook.Worksheets[sheet_test];
            List<string> l = new List<string>();
            string string_number_in_excel;
            string string_step_in_excel;
            string string_spec_min_in_excel;
            string string_spec_max_in_excel;
            int row_in_excel = 1;
            while (true) {
                row_in_excel++;
                if (worksheet.Range[row_in_excel, 1].Style.KnownColor.ToString() == "Color5" || worksheet.Range[row_in_excel, 1].Style.KnownColor.ToString() == "Color2") {
                    continue;
                }
                string_number_in_excel = worksheet.GetText(row_in_excel, 1);
                string_step_in_excel = worksheet.GetText(row_in_excel, 2);
                string_spec_min_in_excel = worksheet.GetNumber(row_in_excel, 3).ToString();
                if (string_spec_min_in_excel == "NaN") string_spec_min_in_excel = worksheet.GetFormulaNumberValue(row_in_excel, 3).ToString();
                if (string_spec_min_in_excel == "NaN") string_spec_min_in_excel = worksheet.GetText(row_in_excel, 3);
                if (string_spec_min_in_excel == null) string_spec_min_in_excel = worksheet.GetFormulaStringValue(row_in_excel, 3);
                if (string_spec_min_in_excel != null) string_spec_min_in_excel = string_spec_min_in_excel.Trim();
                string_spec_max_in_excel = worksheet.GetNumber(row_in_excel, 4).ToString();
                if (string_spec_max_in_excel == "NaN") string_spec_max_in_excel = worksheet.GetFormulaNumberValue(row_in_excel, 4).ToString();
                if (string_spec_max_in_excel == "NaN") string_spec_max_in_excel = worksheet.GetText(row_in_excel, 4);
                if (string_spec_max_in_excel == null) string_spec_max_in_excel = worksheet.GetFormulaStringValue(row_in_excel, 4);
                if (string_spec_max_in_excel != null) string_spec_max_in_excel = string_spec_max_in_excel.Trim();
                if (string_number_in_excel == null) {
                    string_number_in_excel = worksheet.GetNumber(row_in_excel, 1).ToString();
                    if (string_number_in_excel == "NaN") {
                        break;
                    }
                }
                l.Add(string_number_in_excel + ";" + string_step_in_excel + ";" + string_spec_min_in_excel + ";" + string_spec_max_in_excel);
            }
            for (int i = 0; i < s_head.Count(); i++) {
                dataGridView1.Rows.Add(1);
                dataGridView1.Rows[i].Cells[0].Value = i + 1;
                dataGridView1.Rows[i].Cells[1].Value = s_head[i];
                dataGridView1.Rows[i].Cells[2].Value = s_step[i];
                string[] step = s_step[i].Split(',');
                if (step.Count() == 1) {
                    bool hh = false;
                    if (step[0] == "data_fail") { hh = true; dataGridView1.Rows[i].Cells[3].Value = "ok data_fail"; }
                    if (step[0] == "data_Final_Result") { hh = true; dataGridView1.Rows[i].Cells[3].Value = "ok data_Final_Result"; }
                    if (step[0] == "data_DATE_TIME") { hh = true; dataGridView1.Rows[i].Cells[3].Value = "ok data_DATE_TIME"; }
                    if (step[0] == "data_FG") { hh = true; dataGridView1.Rows[i].Cells[3].Value = "ok data_FG"; }
                    if (step[0] == "data_WO") { hh = true; dataGridView1.Rows[i].Cells[3].Value = "ok data_WO"; }
                    if (step[0] == "data_TESTER_ID") { hh = true; dataGridView1.Rows[i].Cells[3].Value = "ok data_TESTER_ID"; }
                    if (step[0] == "data_prism_number") { hh = true; dataGridView1.Rows[i].Cells[3].Value = "ok data_prism_number"; }
                    if (step[0] == "data_Operator") { hh = true; dataGridView1.Rows[i].Cells[3].Value = "ok data_Operator"; }
                    if (step[0] == "data_mode") { hh = true; dataGridView1.Rows[i].Cells[3].Value = "ok data_mode"; }
                    if (step[0] == "data_Test_Start_Time") { hh = true; dataGridView1.Rows[i].Cells[3].Value = "ok data_Test_Start_Time"; }
                    if (step[0] == "data_Test_Finish_Time") { hh = true; dataGridView1.Rows[i].Cells[3].Value = "ok data_Test_Finish_Time"; }
                    if (step[0] == "data_Test_Total_Time") { hh = true; dataGridView1.Rows[i].Cells[3].Value = "ok data_Test_Total_Time"; }
                    if (hh) goto hh_;
                    for (int kk = 0; kk < l.Count; kk++) {
                        string[] mmm = l[kk].Split(';');
                        if (mmm[0] != step[0]) continue;
                        dataGridView1.Rows[i].Cells[3].Value = mmm[1];
                        dataGridView1.Rows[i].Cells[4].Value = mmm[2];
                        dataGridView1.Rows[i].Cells[5].Value = mmm[3];
                        l.Remove(l[kk]);
                        break;
                    }
                    hh_:;
                } else {
                    string data = "";
                    string data_min = "";
                    string data_max = "";
                    foreach (string gh in step) {
                        for (int kk = 0; kk < l.Count; kk++) {
                            string[] mmm = l[kk].Split(';');
                            if (mmm[0] != gh) continue;
                            data += mmm[1] + "\n";
                            data_min += mmm[2] + "\n";
                            data_max += mmm[3] + "\n";
                            l.Remove(l[kk]);
                            break;
                        }
                    }
                    data = data.Substring(0, data.Count() - 1);
                    data_min = data_min.Substring(0, data_min.Count() - 1);
                    data_max = data_max.Substring(0, data_max.Count() - 1);
                    dataGridView1.Rows[i].Cells[3].Value = data;
                    dataGridView1.Rows[i].Cells[4].Value = data_min;
                    dataGridView1.Rows[i].Cells[5].Value = data_max;
                }
            }
            for (int i = 0; i < l.Count; i++) {
                string[] s = l[i].Split(';');
                dataGridView2.Rows.Add(1);
                dataGridView2.Rows[i].Cells[2].Value = s[0];
                dataGridView2.Rows[i].Cells[3].Value = s[1];
                dataGridView2.Rows[i].Cells[4].Value = s[2];
                try { dataGridView2.Rows[i].Cells[5].Value = s[3]; } catch { }
                dataGridView2.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gold;
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e) {
            for (int i = 0; i < dataGridView1.RowCount; i++) {
                dataGridView1.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
            }
            dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.CornflowerBlue;
        }
    }
}
