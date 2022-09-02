using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;

namespace WindowsFormsApp4
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\\Склад.mdb; Jet OLEDB:Database Password=mersedes-0516;");
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter("Select Логин From Администратор where Логин ='" + textBox1.Text + "' and Пароль ='" + textBox2.Text + "'", con);
            DataTable dt = new DataTable();
            dataAdapter.Fill(dt);
            // Проверяем, что количество строк из БД больше нуля
            if (dt.Rows.Count > 0)
            {
                // Нужный Вам ID
                string ID = dt.Rows[0][0].ToString();
                this.Hide();
                Form1 ss = new Form1();
                ss.Show();
            }
            else
            {
                MessageBox.Show("Неправильно введённые Логин или пароль");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form2 form2 = new Form2();
            form2.Show();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox2.UseSystemPasswordChar = true;
        }

        private void Form4_Load(object sender, EventArgs e)
        {

        }
    }
}
