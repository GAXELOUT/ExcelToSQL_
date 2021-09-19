using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace ExcelToSQL.Properties
{
    public partial class Settings_edit : Form
    {
        private Dictionary<string, List<string>> null_table = new Dictionary<string, List<string>> { };
        string code = "";
        System.Collections.Specialized.StringCollection folder_string = Properties.Settings.Default.Path_to_folder_library;
        private bool big_size = false;
        public Settings_edit(string er = "")
        {
            InitializeComponent();
            code = er;
            if(code == "ERORR")
            {
                tabControl1.TabPages.Remove(tabControl1.TabPages[2]);
            }
            button6.Location = new Point(button6.Location.X, button6.Location.Y + 200);
            button3.Location = new Point(button3.Location.X, button3.Location.Y + 200);
            label1.Location = new Point(label1.Location.X, label1.Location.Y + 200);
            label2.Location = label1.Location;
            button4.Location = new Point(button4.Location.X, button4.Location.Y + 200);
            button5.Location = new Point(button5.Location.X, button5.Location.Y + 200);
            dataGridView1.Size = new Size(dataGridView1.Size.Width, dataGridView1.Size.Height + 200);
            dataGridView2.Size = new Size(dataGridView2.Size.Width, dataGridView2.Size.Height + 200);
        }
        private void Lock_load()
        {
            if ((label1.Text == "") || (label2.Text == ""))
            {
                label1.Text = "Подождите, операция выполняется";
                label2.Text = "Подождите, операция выполняется";
                //dataGridView2.Visible = false;
                //dataGridView1.Visible = false;
            }
            else
            {
                label1.Text = "";
                label2.Text = "";
                //dataGridView2.Visible = true;
                //dataGridView1.Visible = true;
            }
            this.Update();
        }
        private void Settings_edit_Load(object sender, EventArgs e)
        {
            SQL_string_name();
            this.CenterToScreen();
            folder_string = Properties.Settings.Default.Path_to_folder_library;
            foreach (string a in Properties.Settings.Default.Path_to_folder_library)
            {
                textBox3.AppendText(a + "\n");
            }
            textBox3.Text = textBox3.Text.Trim();
            textBox2.Text = Search_string(Properties.Settings.Default.SQL_string, "Data Source =").Trim();
            textBox4.Text = Search_string(Properties.Settings.Default.SQL_string);
            //textBox2.Text = Properties.Settings.Default.DB_name;
            textBox1.Text = Properties.Settings.Default.Copy_from;
            //richTextBox4.Text = Properties.Settings.Default.Copy_to;
        }
        private string Search_string(string SQL_st, string st_ser = "Catalog=")
        {
            StringBuilder sb = new StringBuilder();

            for (int i = SQL_st.LastIndexOf(st_ser) + st_ser.Length; i < SQL_st.Length; i++)
            {
                if (SQL_st[i] == ';') 
                {
                    break;
                }
                else
                {
                    sb.Append(SQL_st[i]);
                }
            }

            return sb.ToString();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Сохранить данные?", "", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                try
                {
                    Properties.Settings.Default.Path_to_folder_library.Clear();
                    foreach(string a in textBox3.Text.Split('\n'))
                    {
                        if (a != "")
                        {
                            Properties.Settings.Default.Path_to_folder_library.Add(a);
                        }
                    }
                    Properties.Settings.Default.Copy_from = textBox1.Text;
                    //Properties.Settings.Default.Copy_to = richTextBox4.Text;
                    Properties.Settings.Default.Save();

                    Properties.Settings.Default.SQL_string = @"Data Source=" + @textBox2.Text.Trim()
                        + @";Initial Catalog=" + @textBox4.Text.Trim() + @";Integrated Security=True";
                    Properties.Settings.Default.DB_name = textBox4.Text;
                    Properties.Settings.Default.Save();

                    Mes_from mf = new Mes_from("EDIT");
                    mf.Show();
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Не удалось поменять пути \nОшибка : \n" + ex.Message);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Сохранить данные?", "", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                try
                {
                    //Data Source=NOTEPC\SQLEXPRESS;Initial Catalog=praktikum;Integrated Security=True
                    Properties.Settings.Default.SQL_string = @"Data Source=" + @textBox2.Text.Trim() 
                        + @";Initial Catalog=" + @textBox4.Text.Trim() + @";Integrated Security=True";
                    Properties.Settings.Default.DB_name = textBox4.Text;
                    Properties.Settings.Default.Save();

                    Properties.Settings.Default.Path_to_folder_library.Clear();
                    foreach (string a in textBox3.Text.Split('\n'))
                    {
                        if (a != "")
                        {
                            Properties.Settings.Default.Path_to_folder_library.Add(a);
                        }
                    }
                    Properties.Settings.Default.Copy_from = textBox1.Text;
                    //Properties.Settings.Default.Copy_to = richTextBox4.Text;
                    Properties.Settings.Default.Save();

                    Mes_from mf = new Mes_from("EDIT");
                    mf.ShowDialog();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Не удалось поменять данные \nОшибка : \n" + ex.Message);
                }
            }
        }

        private void Settings_edit_FormClosed(object sender, FormClosedEventArgs e)
        {
            //ExcelToSQL.Form1.ActiveForm.Show();
            Form1.settings_edit = true;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = @"C:\";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                textBox1.Text = dialog.FileName;
            }
        }
        //private void button3_Click(object sender, EventArgs e)
        //{
        //    if (this.MdiChildren.Length != 0)
        //    {
        //        MessageBox.Show("Уже открыта");
        //    }
        //    else
        //    {
        //        Table_create tb = new Table_create(this.Location.X + 591, this.Location.Y);
        //        tb.FormClosed += (o, es) => { Normal_size(); };
        //        this.CenterToScreen();
        //        Size sz_tc = tabControl1.Size;
        //        Point loc_tc = tabControl1.Location;
        //        tabControl1.Dock = DockStyle.None;
        //        tabControl1.Size = sz_tc;
        //        tabControl1.Location = loc_tc;
        //        this.Location = new Point(this.Location.X - 200, this.Location.Y);
        //        //this.Location = new Point(resolution.Width / 2, resolution.Height / 2);
        //        tb.Location = new Point(tabControl1.Location.X + tabControl1.Size.Width, tabControl1.Location.Y );
        //        //tb.Size = new Size
        //        this.IsMdiContainer = true;
        //        tb.Dock = DockStyle.Right;
        //        tb.MdiParent = this;
        //        tb.Show();
        //        this.Size = new Size(tabControl1.Size.Width + tb.Width + 30, tb.Height + 45);

        //    }
        //}

        private void Settings_edit_Activated(object sender, EventArgs e)
        {
            //this.CenterToScreen();
        }
        public void Normal_size()
        {
            this.Size = new Size(550,230);
            tabControl1.Dock = DockStyle.Fill;
            this.CenterToScreen();
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
            dataGridView1.Columns[0].Width = 420;
            dataGridView2.Columns[0].Width = 420;
            dataGridView1[1, 0].Style.BackColor = Color.FromArgb(93, 118, 203);
            dataGridView1[0, 0].Style.BackColor = Color.FromArgb(93, 118, 203);
            dataGridView2[0, 0].Style.BackColor = Color.FromArgb(93, 118, 203);
            dataGridView2[1, 0].Style.BackColor = Color.FromArgb(93, 118, 203);
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
                dataGridView2.Rows.Add("В базе данных не существует таблиц");
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

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 2)
            {
                Load_data();
                this.Size = new Size(this.Size.Width, this.Size.Height + 200);
                big_size = true;
                this.CenterToScreen();
            }
           else
            {
                if(big_size)
                {
                    this.Size = new Size(this.Size.Width, this.Size.Height - 200);
                    big_size = false;
                }
                this.CenterToScreen();
            }
        }
        private List<string> check_list(DataGridView dgv)
        {
            List<string> ls = new List<string>();
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                if ((bool)dgv.Rows[i].Cells[1].EditedFormattedValue)
                {
                    ls.Add(dgv.Rows[i].Cells[0].Value.ToString());
                }
            }
            return ls;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Lock_load();
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
                foreach (var j in compleat_table)
                {
                    null_table.Remove(j);
                }
                //this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            Load_data();
            Lock_load();
            label1.Text = "Было созданно : " + compleat_table.Count.ToString() + " таблиц";
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Lock_load();
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
            Lock_load();
            label2.Text = "Успешно было удалено " + jo.Count + " таблиц";
        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Update();
            dataGridView1.Update();
            dataGridView2.Update();
            label1.Text = "";
            label2.Text = "";
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

        private void button4_Click(object sender, EventArgs e)
        {
            this.Update();
            dataGridView2.Update();

        }
        private List<string> SQL_string_name()
        {
            var sql = Properties.Settings.Default.SQL_string;
            List<string> ls = new List<string>();
            ls.Add(Search_string(sql, "Data Source ="));
            ls.Add(Search_string(sql));
            return ls;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = @"C:\";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                textBox3.Text = dialog.FileName;
            }
        }

        private void Settings_edit_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Opacity = 0;
            this.Update();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            Load_data();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Load_data();
        }
    }
}
