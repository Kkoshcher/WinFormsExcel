using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace WindowsFormsApp1 {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e) {
            comboBox2.Visible = false;
            textBox2.Visible = false;
            comboBox3.Visible = false;
            textBox3.Visible = false;
            comboBox4.Visible = false;
            textBox4.Visible = false;
        }
        private string FormatCellValue(IXLCell cell) {
            if (cell.DataType == XLDataType.DateTime) {
                return cell.GetDateTime().ToString("yyyy/MM/dd");
            }
            return cell.Value.ToString();
        }
        private DataTable originalDataTable;
        private void btnLoad_Click(object sender, EventArgs e) {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK) {
                string filePath = openFileDialog.FileName;
                using (var workbook = new XLWorkbook(filePath)) {
                    var worksheet = workbook.Worksheet(1);
                    DataTable dt = new DataTable();

                    // 清除舊的 ComboBox 項目
                    RepComBtn();

                    // 添加列標題到 DataTable 和 ComboBox
                    var firstRow = worksheet.FirstRowUsed();
                    int columnCount = firstRow.CellsUsed().Count(); 
                    foreach (var cell in firstRow.Cells(1, columnCount)) { 
                        string columnName = cell.Value.ToString();
                        dt.Columns.Add(columnName);
                        comboBox1.Items.Add(columnName);
                        comboBox2.Items.Add(columnName);
                        comboBox3.Items.Add(columnName);
                        comboBox4.Items.Add(columnName);
                    }

                    // 並添加數據
                    var rows = worksheet.RowsUsed().Skip(1);
                    foreach (var row in rows) {
                        dt.Rows.Add(row.Cells(1, columnCount).Select(c => FormatCellValue(c)).ToArray()); 
                    }

                    dataGridView1.DataSource = dt;
                    MessageBox.Show("加載操作：原始數據表行數 = " + dt.Rows.Count);
                    originalDataTable = dt;
                }
            }

        }

        private string FormatCellValue(object value) {
            if (value is DateTime) {
                return ((DateTime)value).ToString("yyyy/MM/dd");
            }
            return value?.ToString() ?? "";
        }

        private void btnDownload_Click(object sender, EventArgs e) {
            DataTable dt = dataGridView1.DataSource as DataTable;
            if (dt == null) {
                DataView dataView = dataGridView1.DataSource as DataView;
                if (dataView != null) {
                    dt = dataView.ToTable();
                }
            }

            if (dt != null) {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Files|*.xlsx";
                saveFileDialog.DefaultExt = "xlsx";
                saveFileDialog.AddExtension = true;

                if (saveFileDialog.ShowDialog() == DialogResult.OK) {
                    using (var workbook = new XLWorkbook()) {
                        var worksheet = workbook.Worksheets.Add("Exported Data");

                        // 添加列標題
                        for (int i = 0; i < dt.Columns.Count; i++) {
                            worksheet.Cell(1, i + 1).Value = dt.Columns[i].ColumnName;
                        }

                        // 添加數據
                        for (int i = 0; i < dt.Rows.Count; i++) {
                            for (int j = 0; j < dt.Columns.Count; j++) {
                                worksheet.Cell(i + 2, j + 1).Value = FormatCellValue(dt.Rows[i][j]);
                            }
                        }

                        workbook.SaveAs(saveFileDialog.FileName);

                        MessageBox.Show("文件已成功保存！");
                        this.Close(); 
                    }
                }
            }
            else {
                MessageBox.Show("無可導出的數據");
            }
        }


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e) {

        }

        private void textBox1_TextChanged(object sender, EventArgs e) {

        }

        private void textBox2_TextChanged(object sender, EventArgs e) {

        }

        private void textBox3_TextChanged(object sender, EventArgs e) {

        }

        private void textBox4_TextChanged(object sender, EventArgs e) {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e) {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e) {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e) {

        }

        private void Selbtn_Click(object sender, EventArgs e) {
            if (originalDataTable != null) {
                var dataView = originalDataTable.DefaultView;

                // 構建篩選條件字符串
                var filterConditions = new List<string>();
                if (comboBox1.SelectedIndex != -1 && !string.IsNullOrEmpty(textBox1.Text)) {
                    filterConditions.Add($"[{comboBox1.SelectedItem.ToString()}] LIKE '%{textBox1.Text}%'");
                }
                if (comboBox2.SelectedIndex != -1 && !string.IsNullOrEmpty(textBox2.Text)) {
                    filterConditions.Add($"[{comboBox2.SelectedItem.ToString()}] LIKE '%{textBox2.Text}%'");
                }
                if (comboBox3.SelectedIndex != -1 && !string.IsNullOrEmpty(textBox3.Text)) {
                    filterConditions.Add($"[{comboBox3.SelectedItem.ToString()}] LIKE '%{textBox3.Text}%'");
                }
                if (comboBox4.SelectedIndex != -1 && !string.IsNullOrEmpty(textBox4.Text)) {
                    filterConditions.Add($"[{comboBox4.SelectedItem.ToString()}] LIKE '%{textBox4.Text}%'");
                }

                // 檢查第一組控件是否已被使用，如果是，則顯示第二組控件
                if (comboBox1.SelectedIndex != -1 && !string.IsNullOrEmpty(textBox1.Text)) {
                    comboBox2.Visible = true;
                    textBox2.Visible = true;
                }

                // 以此類推，檢查第二組控件是否已被使用，如果是，則顯示第三組控件
                if (comboBox2.SelectedIndex != -1 && !string.IsNullOrEmpty(textBox2.Text)) {
                    comboBox3.Visible = true;
                    textBox3.Visible = true;
                }

                // 類似地，檢查第三組控件
                if (comboBox3.SelectedIndex != -1 && !string.IsNullOrEmpty(textBox3.Text)) {
                    comboBox4.Visible = true;
                    textBox4.Visible = true;
                }

                // 應用篩選
                dataView.RowFilter = string.Join(" AND ", filterConditions);

                // 更新 DataGridView 的數據源
                dataGridView1.DataSource = dataView;
            }
        }

        private void btnre_Click(object sender, EventArgs e) {

            RepComBtn();
            // 重新應用篩選
            ApplyFilter();
        }
        private void RepComBtn() {
            // 清空 ComboBox 和 TextBox
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;

            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();

            comboBox2.Visible = false;
            textBox2.Visible = false;
            comboBox3.Visible = false;
            textBox3.Visible = false;
            comboBox4.Visible = false;
            textBox4.Visible = false;
        }
        private void ApplyFilter() {
            if (originalDataTable != null) {
                var dataView = originalDataTable.DefaultView;

                dataView.RowFilter = "";

                dataGridView1.DataSource = dataView;
            }
        }
    }
}
