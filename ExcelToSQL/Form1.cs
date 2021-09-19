using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing.Text;
using OfficeOpenXml;
using System.Threading;

namespace ExcelToSQL
{
    public partial class Form1 : Form
    { 
        private bool perv = true;
        public static bool settings_edit = false;
        DataTable dt_excel = new DataTable();
        DataTable dt_sql = new DataTable();
        PrivateFontCollection private_fonts = new PrivateFontCollection();
        private int width = 0;
        private int height = 0;
        public Form1()
        {
            this.IsMdiContainer = true;
            this.LayoutMdi(MdiLayout.TileHorizontal);
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized;
            SetDoubleBuffered(dataGridView1, true);
            Full_disigner();
            //materialFlatButton1.ForeColor = Color.White;
            //materialFlatButton2.ForeColor = Color.White;
            //materialFlatButton3.ForeColor = Color.White;

        }

        private void Form1_Load(object sender, EventArgs e)
        {

                vis();
                //ExcelPackage package = new ExcelPackage(new FileInfo(@"C:\\Users\\gaxy\\Desktop\\test\\DataBaseAVECS.xlsx"));
                width = this.Size.Width;
                height = this.Size.Height;
                if (!Database_query.Check_con_string())
                {
                    MessageBox.Show("Ошибка при установке соединения базы данных", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Properties.Settings_edit se = new Properties.Settings_edit("ERORR");
                    se.ShowDialog();
                }
                else
                {
                    try
                    {
                        this.Update();
                        FullUpdateData();
                        DGV_designer(dataGridView1);
                        Compare_DataTable(dt_excel, dt_sql);
                    }
                    catch
                    {
                        MessageBox.Show("Ошибка при установке соединения с excel", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Properties.Settings_edit se = new Properties.Settings_edit("ERORR");
                        se.ShowDialog();
                }
                }
                vis();
            
        }

        private void Full_disigner()
        {
            Load_font();
            Button_disigner(button1);
            Button_disigner(button2);
            Button_disigner(button3);
            this.Size = new System.Drawing.Size(height, width);
            tableLayoutPanel1.BackColor = Color.FromArgb(238, 239, 249);
            panel1.BackColor = Color.FromArgb(229,236,247);
            comboBox1.Font = new Font(private_fonts.Families[0], 14);
            tableLayoutPanel1.RowStyles[0].SizeType = SizeType.Absolute;
            tableLayoutPanel1.RowStyles[0].Height = comboBox1.Height + 5;
            Properties.Settings.Default.open = true;
            label1.Font = new Font(private_fonts.Families[0], 14);
            label2.Font = new Font(private_fonts.Families[0], 14);
            //comboBox1.ma
        }

        private void Button_disigner(Button bt)
        {
            bt.BackColor = Color.Transparent; //прозрачный цвет фона
            bt.BackgroundImageLayout = ImageLayout.Center; //выравниваем её по центру
            bt.FlatStyle = FlatStyle.Flat;
            bt.FlatAppearance.BorderSize = 0;//ширина рамки = 0
            bt.TextImageRelation = TextImageRelation.ImageAboveText; //картинка над текстом
            bt.TabStop = false;//делаем так, что бы при потере фокуса, вокруг кнопки не оставалась черная рамка
            bt.MouseEnter += (s, e) => { bt.BackColor = Color.FromArgb(194, 220, 234); };
            bt.MouseLeave += (s, e) => { bt.BackColor = Color.Transparent; };
            bt.Font = new Font(private_fonts.Families[0], 14);
        }

        private void Load_font()
        {
            using (MemoryStream fontStream = new MemoryStream(Properties.Resources.Roboto_Light))
            {
                // create an unsafe memory block for the font data
                System.IntPtr data = Marshal.AllocCoTaskMem((int)fontStream.Length);
                // create a buffer to read in to
                byte[] fontdata = new byte[fontStream.Length];
                // read the font data from the resource
                fontStream.Read(fontdata, 0, (int)fontStream.Length);
                // copy the bytes to the unsafe memory block
                Marshal.Copy(fontdata, 0, data, (int)fontStream.Length);
                // pass the font to the font collection
                private_fonts.AddMemoryFont(data, (int)fontStream.Length);
                // close the resource stream
                fontStream.Close();
                // free the unsafe memory
                Marshal.FreeCoTaskMem(data);

            }
        }
            private void FullUpdateData()
        {
            Path_load pl = new Path_load();
            ComboLoad();
            Update_dt();
            Database_query dbq = new Database_query();
            dataGridView1.DataSource = dt_excel;
        }

        private void ComboLoad()
        {
            comboBox1.DataSource = new List<string>(Path_load.Path_dic.Keys);
        }

        private void Update_dt()
        {
            Excel_reader er = new Excel_reader();
            Database_query db_q = new Database_query();
            dt_excel = er.GenerateDataTable(comboBox1.Text);
            dt_sql = db_q.BD_load(comboBox1.Text);
        }
        
        private void DGV_designer(DataGridView dgv)
        {
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgv.RowHeadersVisible = false;
            //dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgv.Dock = DockStyle.Fill;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249);
            dgv.BorderStyle = BorderStyle.None;
            //dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(209,194,240);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            dgv.DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            dgv.EnableHeadersVisualStyles = false;
            foreach (DataGridViewColumn column in dgv.Columns)
            {
                column.Width = 100;
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            dgv.Columns[1].Width = 200;
        }
        private void Compare_DataTable(DataTable dt1_ex, DataTable dt2_sq)
        {
            for (int j = 0; j < dt1_ex.Rows.Count; j++)
            {
                for (int i = 0; i < dt1_ex.Columns.Count; i++)
                {

                    if ((dt2_sq.Rows.Count > j) && (dt2_sq.Columns.Count > i))
                    {
                        if (dt1_ex.Rows[j][i].ToString().Trim() != (dt2_sq.Rows[j][i].ToString().Trim()))
                        {
                            dataGridView1.Rows[j].Cells[i].Style.BackColor = System.Drawing.Color.FromArgb(255, 255, 53);
                        }
                    }
                    else
                    {
                        dataGridView1.Rows[j].Cells[i].Style.BackColor = System.Drawing.Color.FromArgb(61, 177, 68);
                    }
                }
            }
        }
        public void SetDoubleBuffered(Control c, bool value)
        {
            PropertyInfo pi = typeof(Control).GetProperty("DoubleBuffered", BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.NonPublic);
            if (pi != null)
            {
                pi.SetValue(c, value, null);

                MethodInfo mi = typeof(Control).GetMethod("SetStyle", BindingFlags.Instance | BindingFlags.InvokeMethod | BindingFlags.NonPublic);
                if (mi != null)
                {
                    mi.Invoke(c, new object[] { ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true });
                }

                mi = typeof(Control).GetMethod("UpdateStyles", BindingFlags.Instance | BindingFlags.InvokeMethod | BindingFlags.NonPublic);
                if (mi != null)
                {
                    mi.Invoke(c, null);
                }
            }

        }


        //Update_dt();
        //dataGridView1.DataSource = dt_excel;
        //    DGV_designer(dataGridView1);
        //Compare_DataTable(dt_excel, dt_sql);
        //dataGridView1.Focus();
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!perv)
            {
                vis();
                perv = false;
                Update_dt();
                dataGridView1.DataSource = dt_excel;
                Compare_DataTable(dt_excel, dt_sql);
                dataGridView1.Focus();
                vis();
                DGV_designer(dataGridView1);
            }
            else
            {
                perv = false;
            }
        }
       
        private void comboBox1_DropDownClosed(object sender, EventArgs e)
        {
            dataGridView1.Focus();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Database_query bdl = new Database_query();
            vis();
            if (dt_excel.Columns.Count != dt_sql.Columns.Count)
            {
                bdl.RE_Creating_table(dt_excel, comboBox1.Text);
                goto M;
            }
            else
            {
                for(int i = 0; i < dt_excel.Columns.Count; i++ )
                {
                    if (dt_excel.Columns[i].ColumnName != dt_sql.Columns[i].ColumnName)
                    {
                        bdl.RE_Creating_table(dt_excel, comboBox1.Text);
                        goto M;
                    }
                }
            }

                bdl.Insert_query(dt_excel, comboBox1.Text);
                M:
                Update_dt();
                dataGridView1.DataSource = dt_excel;
                Compare_DataTable(dt_excel, dt_sql);
                dataGridView1.Focus();
            vis();
            DGV_designer(dataGridView1);
        }
        private void vis()
        {
            this.Enabled = !this.Enabled;
            label1.Visible = !label1.Visible;
            //dataGridView1.Visible = !dataGridView1.Visible;
            if((!label1.Visible)&& (Path_load.no_name_table.Count != 0))
            {
                label2.Visible = true;
            }
            else
            {
                label2.Visible = false;
            }
            this.Update();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            vis();
            Update_dt();
            dataGridView1.DataSource = dt_excel;
            Compare_DataTable(dt_excel, dt_sql);
            dataGridView1.Focus();
            vis();
            DGV_designer(dataGridView1);
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            this.Size = new System.Drawing.Size(height, width);
            if (settings_edit)
            {
                vis();
                if (!Database_query.Check_con_string())
                {
                    MessageBox.Show("Ошибка при установке соединения базы данных", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Properties.Settings_edit se = new Properties.Settings_edit("ERORR");
                    se.ShowDialog();
                    Application.Exit();

                }
                else
                {
                    try
                    {
                        perv = true;
                        Full_disigner();
                        this.Update();
                        FullUpdateData();
                        DGV_designer(dataGridView1);
                        Compare_DataTable(dt_excel, dt_sql);
                    }
                    catch 
                    {
                        MessageBox.Show("Ошибка при установке соединения с excel", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Properties.Settings_edit se = new Properties.Settings_edit("ERORR");
                        se.ShowDialog();
                        Application.Exit();
                    }
                }
                vis();
                settings_edit = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Properties.Settings_edit se = new Properties.Settings_edit("");
            se.ShowDialog();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if ((Properties.Settings.Default.Copy_to != "") && (Properties.Settings.Default.Copy_from != ""))
            {
                DialogResult res = MessageBox.Show("Вы хотите произвести резервное копирование директории на локальный диск  ?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    this.Hide();
                    Copy_directory cd = new Copy_directory();
                    cd.ShowDialog();
                }
            }
        }

        private void dataGridView1_VisibleChanged(object sender, EventArgs e)
        {
            //materialLabel1.Visible = !materialLabel1.Visible;

        }

        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView1.Visible = false;
            Point cen = new Point(Screen.PrimaryScreen.Bounds.Height, Screen.PrimaryScreen.Bounds.Width);
            Test1 t1 = new Test1(cen.X - 200,cen.Y);
            Test2 t2 = new Test2(cen.X + 200,cen.Y);
            t1.MdiParent = this;
            t2.MdiParent = this;
            this.Opacity = 100;
            this.Update();
            t1.Show();
            t2.Show();
        }
    }
}
