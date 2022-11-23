using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace zadacha
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void Form2_Load(object sender, EventArgs e)
        {
            Form1 main = this.Owner as Form1;
            if (main != null)
            {
                dateTimePicker1.Value = main.dateTimePicker1.Value;

                textBox1.Text = main.textBox1.Text;
                textBox2.Text = main.textBox2.Text;
                textBox3.Text = main.textBox3.Text;
                textBox4.Text = main.textBox4.Text;
                textBox5.Text = main.textBox5.Text;
                textBox6.Text = main.textBox6.Text;
                textBox7.Text = main.textBox7.Text;

                dataGridView1.Rows[0].Cells[0].Value = main.comboBox1.Text;
                dataGridView1.Rows[0].Cells[1].Value = main.comboBox2.Text;
                dataGridView1.Rows[0].Cells[2].Value = main.comboBox3.Text;
                dataGridView1.Rows[0].Cells[3].Value = main.comboBox4.Text;
                dataGridView1.Rows[0].Cells[4].Value = main.textBox1.Text;
                dataGridView1.Rows[0].Cells[5].Value = main.textBox2.Text;
                dataGridView1.Rows[0].Cells[6].Value = main.textBox4.Text;
                dataGridView1.Rows[0].Cells[7].Value = main.textBox3.Text;
                dataGridView1.Rows[0].Cells[8].Value = main.dateTimePicker1.Value.ToString();
                dataGridView1.Rows[0].Cells[9].Value = main.textBox5.Text;
                dataGridView1.Rows[0].Cells[10].Value = main.textBox7.Text;
                dataGridView1.Rows[0].Cells[11].Value = main.textBox6.Text;
            }
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox9.SelectedIndex = comboBox8.SelectedIndex;

            if (comboBox8.Text == "РОССИЯ")
            {
                maskedTextBox1.Visible = true;
                label23.Visible = true;
            }
            else
            {
                maskedTextBox1.Visible = false;
                label23.Visible = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form3 ff = new Form3();
            ff.Owner = this;
            ff.Show();
            this.Hide();
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

        private void label6_Click(object sender, EventArgs e)
        {

        }
    }
}
