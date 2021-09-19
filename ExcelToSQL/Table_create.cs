using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToSQL
{
    public partial class Table_create : Form
    {
        private Dictionary<string, List<string>> null_table = new Dictionary<string, List<string>> { };
        public Table_create(int x = 0, int y = 0, Dictionary<string, List<string>> table = null)
        {
            if (x + y != 0)
            {
                this.Location = new Point(x,y);
            }
            else
            {
                this.CenterToScreen();
            }
            null_table = table;
            InitializeComponent();
        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {
           
        }

        private void Table_create_Load(object sender, EventArgs e)
        {

            if (null_table == null)
            {
                Load_data();
            }
            else
            {
                Load_data(false);
            }
        }
        private void Load_data(bool Load = true)
        {
            if (Load)
            {
                null_table = Path_load.tab_update();
            }
            label1.Text = "";
            label2.Text = "";
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            dataGridView1.Rows.Add("Все");
            dataGridView2.Rows.Add("Все");
            dataGridView1[0, 0].Style.BackColor = Color.FromArgb(93, 118, 203);
            foreach (var a in null_table.Keys)
            {
                dataGridView1.Rows.Add(a);
            }
            List<string> s = Database_query.Tab_name();
            foreach (var f in s)
            {
                dataGridView2.Rows.Add(f);
            }


            if (dataGridView2.Rows.Count == 1)
            {
                dataGridView2.Rows.Clear();
                dataGridView2.Columns[1].Visible = false;
                dataGridView2.Rows.Add("Таблица пуста");
            }
            else
            {
                dataGridView2.Columns[1].Visible = true;
            }


            if (dataGridView1.Rows.Count == 1)
            {
                dataGridView1.Rows.Clear();
                dataGridView1.Columns[1].Visible = false;
                dataGridView1.Rows.Add("Все таблицы уже существуют");
            }
            else
            {
                dataGridView1.Columns[1].Visible = true;
            }

            dataGridView1.CurrentCell = null;//Изменяем CurrentCell
            dataGridView2.CurrentCell = null;//Изменяем CurrentCell

            this.Update();
        }

        //private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //   
        //}
        private List<string> check_list(DataGridView dgv)
        {
            List<string> ls = new List<string>();
            for(int i = 0; i<dgv.Rows.Count;i++)
            {
                if ((bool)dgv.Rows[i].Cells[1].EditedFormattedValue)
                {
                    ls.Add(dgv.Rows[i].Cells[0].Value.ToString());
                }
            }
            return ls;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            block();
            List<string> compleat_table = new List<string>();
            try
            {
                var check_list_table = check_list(dataGridView1);
                List<string> selected_null_table = new List<string>();
                selected_null_table.AddRange(from string a in check_list_table
                                             where (a != "Все")
                                             select a);
                foreach (string c in selected_null_table)
                {
                    try
                    {
                        Database_query.Create_table(c, null_table[c]);
                        compleat_table.Add(c);
                    }
                    catch (Exception ex)
                    {
                        compleat_table.Add(c);
                        MessageBox.Show("Ошибка при создании таблицы : " + c + "\nТекст ошибки :\n" + ex.Message);
                    }
                }
                foreach(var j in compleat_table)
                {
                    null_table.Remove(j);
                }
                //this.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            Load_data();
            block();
            label1.Text = "Было созданно : " + compleat_table.Count.ToString() + " таблиц";
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if ((e.RowIndex == 0) && (e.ColumnIndex == 1))
            {
                if ((bool)dataGridView1[1, 0].EditedFormattedValue)
                {
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        dataGridView1[1, i].Value = (bool)dataGridView1[1, 0].EditedFormattedValue;
                    }
                }
                else
                {
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        dataGridView1[1, i].Value = false;
                    }
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void block()
        {
            if ((label1.Text == "") || (label2.Text == ""))
            {
                label1.Text = "Подождите, операция выполняется";
                label2.Text = "Подождите, операция выполняется";
            }
            else
            {
                label1.Text = "";
                label2.Text = "";
            }
            this.Update();
        }
        private void button3_Click_1(object sender, EventArgs e)
        {
            block();
            var check_list_table = check_list(dataGridView2);
            List<string> selected_null_table = new List<string>();
            selected_null_table.AddRange(from string a in check_list_table
                                         where (a != "Все")
                                         select a);
            List<string> jo = new List<string>();

            foreach (string a in selected_null_table)
            {
                Database_query.delete_table(a);
                jo.Add(a);
            }
            Load_data();
            block();
            label2.Text = "Успешно было удалено " + jo.Count + " таблиц";

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.CurrentCell = null;//Изменяем CurrentCell
            if ((e.RowIndex == 0) && (e.ColumnIndex == 1))
            {
                if ((bool)dataGridView2[1, 0].EditedFormattedValue)
                {
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        dataGridView2[1, i].Value = (bool)dataGridView2[1, 0].EditedFormattedValue;
                    }
                }
                else
                {
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        dataGridView2[1, i].Value = false;
                    }
                }
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            label1.Text = "";
            label2.Text = "";
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Table_create_Leave(object sender, EventArgs e)
        {
            this.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Table_create_FormClosing(object sender, FormClosingEventArgs e)
        {
            
        }

        private void Table_create_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (this.Parent != null)
            {
                Parent.Size = new Size(tabControl1.Size.Width, tabControl1.Size.Height);
                this.CenterToScreen();
                
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Load_data();
        }

        //    private void button1_Click(object sender, EventArgs e)
        //    {
        //       
        //    }
    }
}
