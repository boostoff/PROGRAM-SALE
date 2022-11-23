using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.IO;
using System.Diagnostics;

namespace zadacha
{
    public partial class Form3 : Form
    {
        private object _missingObj = System.Reflection.Missing.Value;
        public Form3()
        {
            InitializeComponent();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Нет")
            {
                button1.Enabled = true;
                this.Size = new System.Drawing.Size(690, 150);
            }

            if (comboBox1.Text == "Да")
            {
                button1.Enabled = true;
                this.Size = new System.Drawing.Size(690, 615);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = Directory.GetCurrentDirectory();
            Excel.Application ExcelApp = new Excel.Application();
            Excel.Workbook workbook;
            workbook = ExcelApp.Workbooks.Open(path + (@"\Shablon.xlsx"));
            Excel.Worksheet sheet = (Excel.Worksheet)ExcelApp.ActiveSheet;
            Excel.Range range1 = sheet.get_Range("A1", "AT500");
            Excel.Range range2 = sheet.get_Range("AM1", "AT500");

            sheet.Columns.ClearFormats();
            sheet.Rows.ClearFormats();


            int lastRowIgnoreFormulas = sheet.Cells.Find(
                            "*",
                            System.Reflection.Missing.Value,
                            Excel.XlFindLookIn.xlValues,
                            Excel.XlLookAt.xlWhole,
                            Excel.XlSearchOrder.xlByRows,
                            Excel.XlSearchDirection.xlPrevious,
                            false,
                            System.Reflection.Missing.Value,
                            System.Reflection.Missing.Value).Row;

            int lastColIgnoreFormulas = sheet.Cells.Find(
                            "*",
                            System.Reflection.Missing.Value,
                            System.Reflection.Missing.Value,
                            System.Reflection.Missing.Value,
                            Excel.XlSearchOrder.xlByColumns,
                            Excel.XlSearchDirection.xlPrevious,
                            false,
                            System.Reflection.Missing.Value,
                            System.Reflection.Missing.Value).Column;

            int lastColIncludeFormulas = sheet.UsedRange.Columns.Count;
            int lastRowIncludeFormulas = sheet.UsedRange.Rows.Count;

            int i;
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                sheet.Cells[lastRowIgnoreFormulas + 1, i + 1] = dataGridView1.Rows[0].Cells[i].Value;
            }

            sheet.Rows.RowHeight = 70;
            sheet.Columns.ColumnWidth = 25;
            range1.Cells.Font.Name = "Times New Roman";
            range1.Cells.Font.Size = 14;
            range1.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
            range1.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
            range1.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
            range1.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
            range1.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
            range1.WrapText = true;
            range1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            ExcelApp.Visible = true;


        }

        private void Form3_Load(object sender, EventArgs e)
        {
            Form2 main = this.Owner as Form2;
            if (main != null)
            {
                dataGridView1.Rows[0].Cells[0].Value = main.textBox1.Text;
                dataGridView1.Rows[0].Cells[1].Value = main.comboBox1.Text;
                dataGridView1.Rows[0].Cells[6].Value = main.comboBox3.Text;
                dataGridView1.Rows[0].Cells[7].Value = main.textBox2.Text;
                dataGridView1.Rows[0].Cells[8].Value = main.textBox4.Text;
                dataGridView1.Rows[0].Cells[9].Value = main.dateTimePicker1.Value.ToString();
                dataGridView1.Rows[0].Cells[10].Value = main.textBox3.Text;
                dataGridView1.Rows[0].Cells[11].Value = main.comboBox4.Text;
                dataGridView1.Rows[0].Cells[12].Value = main.comboBox6.Text;
                dataGridView1.Rows[0].Cells[13].Value = main.comboBox6.Text;
                dataGridView1.Rows[0].Cells[14].Value = main.comboBox7.Text;
                dataGridView1.Rows[0].Cells[15].Value = main.textBox8.Text;
                dataGridView1.Rows[0].Cells[16].Value = main.textBox9.Text;
                dataGridView1.Rows[0].Cells[17].Value = main.comboBox5.Text;
                dataGridView1.Rows[0].Cells[18].Value = main.textBox5.Text;
                dataGridView1.Rows[0].Cells[19].Value = main.textBox7.Text;
                dataGridView1.Rows[0].Cells[20].Value = main.textBox6.Text;
                dataGridView1.Rows[0].Cells[21].Value = main.dateTimePicker2.Value.ToString();
                dataGridView1.Rows[0].Cells[22].Value = main.comboBox2.Text;
                dataGridView1.Rows[0].Cells[23].Value = main.maskedTextBox1.Text;
                dataGridView1.Rows[0].Cells[24].Value = main.comboBox9.Text;
                dataGridView1.Rows[0].Cells[25].Value = "Очная";
                dataGridView1.Rows[0].Cells[26].Value = main.comboBox11.Text;
                dataGridView1.Rows[0].Cells[27].Value = main.comboBox10.Text;

                dataGridView1.Rows[0].Cells[2].Value = main.dataGridView1.Rows[0].Cells[0].Value;
                dataGridView1.Rows[0].Cells[3].Value = main.dataGridView1.Rows[0].Cells[1].Value;
                dataGridView1.Rows[0].Cells[4].Value = main.dataGridView1.Rows[0].Cells[2].Value;
                dataGridView1.Rows[0].Cells[5].Value = main.dataGridView1.Rows[0].Cells[3].Value;
                dataGridView1.Rows[0].Cells[28].Value = comboBox1.Text;
                dataGridView1.Rows[0].Cells[29].Value = textBox8.Text;
                dataGridView1.Rows[0].Cells[30].Value = dateTimePicker2.Value.ToString();
                dataGridView1.Rows[0].Cells[31].Value = textBox14.Text; dateTimePicker2.Value.ToString();
                dataGridView1.Rows[0].Cells[32].Value = textBox13.Text;
                dataGridView1.Rows[0].Cells[33].Value = textBox12.Text;
                dataGridView1.Rows[0].Cells[34].Value = textBox9.Text;
                dataGridView1.Rows[0].Cells[35].Value = textBox11.Text;
                dataGridView1.Rows[0].Cells[36].Value = textBox10.Text;
                dataGridView1.Rows[0].Cells[37].Value = comboBox2.Text;
                dataGridView1.Rows[0].Cells[38].Value = main.textBox1.Text;
                dataGridView1.Rows[0].Cells[39].Value = main.textBox2.Text;
                dataGridView1.Rows[0].Cells[40].Value = main.textBox4.Text;
                dataGridView1.Rows[0].Cells[41].Value = main.textBox3.Text;
                dataGridView1.Rows[0].Cells[42].Value = main.dateTimePicker1.Value.ToString();
                dataGridView1.Rows[0].Cells[43].Value = main.textBox5.Text;
                dataGridView1.Rows[0].Cells[44].Value = main.textBox7.Text;
                dataGridView1.Rows[0].Cells[45].Value = main.textBox6.Text;


            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                proc.Kill();
            }

            Application.Exit();
        }
    }   
}
