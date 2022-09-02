using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace WindowsFormsApp4
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }
        public static string BD = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.Combine(Application.StartupPath, "Склад.mdb; Jet OLEDB:Database Password=mersedes-0516;");
        private OleDbConnection con;
        private void Form3_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = this.PopulateDataGridView();

            string put = "SELECT * FROM Товар";
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(put, BD);
            // создаем объект DataSet
            DataSet ds = new DataSet();
            // заполняем таблицу Order  
            // данными из базы данных

            dataAdapter.Fill(ds, "Товар");
            dataGridView1.DataSource = ds.Tables[0].DefaultView;
            dataGridView1.Sort(this.dataGridView1.Columns["код"], ListSortDirection.Ascending);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox6.Text == "")
            {
                MessageBox.Show("Введите данные для поиска");
            }
            else
            {
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                bs.Filter = "[Наименование] like '%" + textBox6.Text + "%' " +
                    "OR [Артикул] like '%" + textBox6.Text + "%'";
                dataGridView1.DataSource = bs;
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox6.Clear();


            string put = "SELECT * FROM Товар";
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(put, BD);
            // создаем объект DataSet
            DataSet ds = new DataSet();
            // заполняем таблицу Order  
            // данными из базы данных

            dataAdapter.Fill(ds, "Товар");
            dataGridView1.DataSource = ds.Tables[0].DefaultView;
            dataGridView1.Sort(this.dataGridView1.Columns["код"], ListSortDirection.Ascending);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form2 form = new Form2();
            form.Show();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox6_KeyUp(object sender, KeyEventArgs e)
        {
            dataGridView1.DataSource = this.PopulateDataGridView();
        }

        private DataTable PopulateDataGridView()
        {
            string query = "SELECT * FROM Товар";
            query += " WHERE Наименование + Артикул  LIKE '%' + @Наименование + '%'";
            query += " OR @Наименование = ''";
            string constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.Combine(Application.StartupPath, "Склад.mdb; Jet OLEDB:Database Password=mersedes-0516;");

            using (OleDbConnection con = new OleDbConnection(constr))
            {
                using (OleDbCommand cmd = new OleDbCommand(query, con))
                {

                    cmd.Parameters.AddWithValue("@Наименование", textBox6.Text.Trim());
                    using (OleDbDataAdapter sda = new OleDbDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        sda.Fill(dt);
                        return dt;


                    }
                }
            }


        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            if (e.RowIndex < 0)
            {
                return;
            }
            int index = e.RowIndex;
        }

        private void главнаяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form2 form2 = new Form2();
            form2.Show();
        }

        private void emexruToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://emex.ru/products/" + textBox1.Text);
        }

        private void existruToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://exist.ru/Price/?pcode=" + textBox1.Text);
        }
    }
}
