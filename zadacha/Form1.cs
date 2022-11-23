using System.Diagnostics;
namespace zadacha
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Оригинал")
            {
                button1.Enabled = true;
                this.Size = new System.Drawing.Size(590, 150);
            }
            
            if (comboBox1.Text == "Дубликат")
            {
                button1.Enabled = true;
                this.Size = new System.Drawing.Size(590, 650);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 f = new Form2();
            f.Owner = this;
            f.Show();
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
    }
}