using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Timers;
using System.IO;
using ClosedXML.Excel;

namespace Export_Data
{
    public partial class Form1 : Form
    {
        private System.Timers.Timer exportTimer;

        private string connectionString = "Data Source=10.4.17.184;Initial Catalog=DEMO;User ID=tk2;Password=123456;Integrated Security=False;";
        private string exportFolder = @"D:\Desktop\Practice\";
        private DataTable currentData;
        public Form1()
        {
            InitializeComponent();
            exportTimer = new System.Timers.Timer();
            exportTimer.Interval = 10 * 1000;
            exportTimer.Elapsed += ExportTimer_Elapsed;
            exportTimer.Start();
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void ExportDataToExcel(DataTable dt)
        {
            try
            {
                // Đếm số file đã export trước đó
                string[] existingFiles = Directory.GetFiles(exportFolder, "Export*.xlsx");
                int nextIndex = existingFiles.Length + 1;

                string fileName = $"Export{nextIndex}.xlsx";
                string filePath = Path.Combine(exportFolder, fileName);

                // Tạo workbook Excel và thêm dữ liệu
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add(dt, "Data");
                    workbook.SaveAs(filePath);
                }

                MessageBox.Show($"Export thành công: {fileName}", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xuất file Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    MessageBox.Show("Kết nối thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kết nối thất bại: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            currentData = FetchDataFromSql();
            dataGridView1.DataSource = currentData;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (currentData != null && currentData.Rows.Count > 0)
            {
                ExportDataToExcel(currentData);
                MessageBox.Show("Đã xuất dữ liệu ra file .xlsx thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Không có dữ liệu để xuất.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void ExportTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            DataTable data = FetchDataFromSql();
            ExportDataToExcel(data);
        }
        private DataTable FetchDataFromSql()
        {
            DataTable dt = new DataTable();

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string query = "SELECT TOP 50 * FROM View_1";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi truy vấn dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return dt;
        }
        private void ExportDataToTxt(DataTable dt)
        {
            try
            {
                // Đếm số file đã export trước đó
                string[] existingFiles = Directory.GetFiles(exportFolder, "Export*.txt");
                int nextIndex = existingFiles.Length + 1;

                string fileName = $"Export{nextIndex}.txt";
                string filePath = Path.Combine(exportFolder, fileName);
                StringBuilder sb = new StringBuilder();

                // Header
                foreach (DataColumn col in dt.Columns)
                {
                    sb.Append(col.ColumnName + "\t");
                }
                sb.AppendLine();

                // Data
                foreach (DataRow row in dt.Rows)
                {
                    foreach (var item in row.ItemArray)
                    {
                        sb.Append(item.ToString() + "\t");
                    }
                    sb.AppendLine();
                }

                File.WriteAllText(filePath, sb.ToString());
                MessageBox.Show($"Export thành công: {fileName}", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xuất file: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void button4_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
