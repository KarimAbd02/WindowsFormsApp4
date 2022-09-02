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
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public static string BD = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.Combine(Application.StartupPath, "Склад.mdb; Jet OLEDB:Database Password=mersedes-0516;");
        private OleDbConnection con;
        
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.ReadOnly = true;
        }
        private void Form1_Load(object sender, EventArgs e)
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

       
        
        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            

            if (e.RowIndex < 0)
            {
                return;
            }
            
                dataGridView1.Columns[0].HeaderCell.Style.Font = new System.Drawing.Font("Microsoft Sans Serif", 14);
            
            int index = e.RowIndex;
            dataGridView1.Rows[index].Selected = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try { 
            con = new OleDbConnection(BD);
            con.Open();
            string queryString = "UPDATE Товар SET [Наименование]  = '" + textBox2.Text + "',[шт]  = '" + textBox3.Text + "',[Артикул] =  = '" + textBox4.Text + "',[Цена] =  = '" + textBox5.Text + "',[Склад] =  = '" + textBox7.Text + "' WHERE [код]=" + textBox1.Text;
            OleDbCommand command = new OleDbCommand(queryString, con);
            command.ExecuteNonQuery();
                dataGridView1.Refresh();
                con.Close();
            MessageBox.Show("Успешно изменен!");
                




                textBox1.Clear();
            textBox3.Clear();
            textBox2.Clear();
            textBox4.Clear();
            textBox5.Clear();
                textBox7.Clear();
            }
            catch
            {
                MessageBox.Show("Выберите строку");
            }
            dataGridView1.Refresh();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
         
            

            if (e.RowIndex < 0)
            {
                return;
            }

            int index = e.RowIndex;
            dataGridView1.Rows[index].Selected = true;

            textBox1.ReadOnly = true;
            textBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox7.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            textBox5.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
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
       

        private void button1_Click(object sender, EventArgs e)
        {
            

           
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            
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





        
        private void button3_Click(object sender, EventArgs e)
        {
            
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
                //Книга.
                ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
                //Таблица.
                ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

                //dataGridView2.Columns[0].HeaderText = "Наименование";
                // dataGridView2.Columns[1].HeaderText = "Артикул";
                // dataGridView2.Columns[2].HeaderText = "Цена";


                dataGridView2.ClearSelection();
                dataGridView2.Columns[0].Width = 195;

                Excel.Application ex;                       // объявляем переменную для приложения
                ex = new Excel.Application();               // связываем переменную с Excel
                ex.Visible = true;                          // отображаем на экране
                Excel.Workbook ex_book = ex.Workbooks.Add(); // создаем в приложении новую рабочую книгу



                ex.Worksheets[1].Range(ex.Worksheets[1].Cells[1, 1], ex.Worksheets[1].Cells[1, 3]).Merge();
                ex.Worksheets[1].Cells[1, 1] = "Накладная " + DateTime.Now.ToShortDateString();
                
            // заполняем автоматически таблицу в Excel данными из таблицы DataGridView1
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView2.Columns.Count; j++)
                    {
                        if (i == 0)
                            ex.Worksheets[1].Cells[i + 2, j + 1] = dataGridView2.Columns[j].HeaderText.ToString();
                        else
                            ex.Worksheets[1].Cells[i + 2, j + 1].Value = dataGridView2.Rows[i - 1].Cells[j].Value.ToString();
                    }
                 
                }

           


            Excel.Range xl_range = ex.Worksheets[1].Range(ex.Worksheets[1].Cells[2, 1],
                    ex.Worksheets[1].Cells[dataGridView2.Rows.Count + 1, dataGridView2.Columns.Count]); // выделение заполненной таблицы в Excel 

                // форматирование выделенного диапазона
                xl_range.Cells.Font.Name = "Tahoma";                      // шрифт диапазона
                xl_range.Cells.Font.Size = 12;                            // размер шрифта для диапазона
                xl_range.Cells.Font.Color = Excel.XlRgbColor.rgbBlack;    // цвет шрифта диапазона
                xl_range.Interior.Color = Excel.XlRgbColor.rgbWhite;  // фоновый цвет диапазона

                // прорисовка границ таблицы выделенного диапазона
                xl_range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                xl_range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                xl_range.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
                xl_range.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
                xl_range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

                xl_range.EntireColumn.AutoFit();      // автоширина и автовысота
                xl_range.EntireRow.AutoFit();

                Excel.Range xl_range1 = ex.Worksheets[1].Range(ex.Worksheets[1].Cells[1, 1], ex.Worksheets[1].Cells[1, 2]);
                // форматирование выделенного диапазона
                xl_range1.Cells.Font.Name = "Candara";                      // шрифт диапазона
                xl_range1.Cells.Font.Size = 20;                            // размер шрифта для диапазона
                xl_range1.Cells.Font.Color = Excel.XlRgbColor.rgbBlack;    // цвет шрифта диапазона
                xl_range1.Interior.Color = Excel.XlRgbColor.rgbWhite;
        }

       


        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
          
        }

        
    

        private void button6_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox4.Text == "" || textBox5.Text == "")
            {
                MessageBox.Show("Выберите поле и внесите: Таблица - нажатие по необходимой строке - Внести в наклдную");
            }
            else
            {

            string firstColum = textBox2.Text;
            string secondColum = textBox4.Text;
            string secondColumn = textBox5.Text;
            string[] row = { firstColum, secondColum, secondColumn };
            dataGridView2.Rows.Add(row);
            dataGridView2.AutoResizeColumns();

            textBox1.Clear();
            textBox3.Clear();
            textBox2.Clear();
            textBox4.Clear();
            textBox5.Clear();
            }
           
        }

        

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView2.AutoResizeColumns();

          

            if (e.RowIndex < 0)
            {
                return;
            }

            int index = e.RowIndex;
            dataGridView2.Rows[index].Selected = true;

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == "" || textBox4.Text == "" || textBox5.Text == "" || textBox7.Text == "" || textBox3.Text == "")
            {
                MessageBox.Show("Заполните все поля");
            }
            else
            {

                OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.Combine(Application.StartupPath, "Склад.mdb;  Jet OLEDB:Database Password=mersedes-0516;"));
                con.Open();
                string queryString = "INSERT INTO Товар ([Наименование],[Склад],[шт],[Артикул],[Цена]) values('" + textBox2.Text + "','" + textBox7.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "')";
                OleDbCommand command = new OleDbCommand(queryString, con);
                command.ExecuteNonQuery();
                this.Refresh();
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter("SELECT * FROM Товар", con);
                DataSet ds = new DataSet();
                dataAdapter.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0].DefaultView;
                dataGridView1.Sort(this.dataGridView1.Columns["код"], ListSortDirection.Ascending);
            }
        }

        private void dataGridView1_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dataGridView1.Sort(this.dataGridView1.Columns["код"], ListSortDirection.Ascending);

            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            
                foreach (DataGridViewRow row in dataGridView2.SelectedRows)
                {
                    dataGridView2.Rows.RemoveAt(row.Index);
                }
            
        }

        private void button8_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form2 form2 = new Form2();
            form2.Show();
        }

        private void textBox6_KeyUp(object sender, KeyEventArgs e)
        {
            dataGridView1.DataSource = this.PopulateDataGridView();

           
        }

        private void emexruToolStripMenuItem_Click(object sender, EventArgs e)
        {
           

        }

        private void главнаяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form2 form2 = new Form2();
            form2.Show();
        }

        private void emexruToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (textBox4.Text == "")
                MessageBox.Show("Артикул не указан. Обратитесь к администратору");
            else
            {
                System.Diagnostics.Process.Start("https://emex.ru/products/" + textBox4.Text);
            }
            
        }

        private void existruToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (textBox4.Text == "")
                MessageBox.Show("Артикул не указан. Обратитесь к администратору");
            else
            {
                System.Diagnostics.Process.Start("https://exist.ru/Price/?pcode=" + textBox4.Text);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox7.Clear();
        }
    }
}
