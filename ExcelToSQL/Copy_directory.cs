using System;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace ExcelToSQL
{
    public partial class Copy_directory : Form
    {
        public static string copy_to = "";
        public Copy_directory()
        {
            InitializeComponent();
            designer();
        }
        private void designer()
        {
            tableLayoutPanel1.RowStyles[0].Height = 100;
            tableLayoutPanel1.RowStyles[1].Height = 0;

            //Дизайн для основного выбора

            this.Size = new Size(550, 170);
            materialLabel1.Location = new Point(-5, 5);
            materialLabel2.Location = new Point(-5, 35);
            richTextBox1.Location = new Point(materialLabel2.Size.Width  , 35);
            richTextBox2.Location = new Point(materialLabel2.Size.Width  , 5);
            richTextBox1.Size = new Size(200,25);
            richTextBox2.Size = new Size(200,25);
            pictureBox1.Location = new Point(richTextBox1.Location.X + richTextBox1.Size.Width + 2, richTextBox1.Location.Y);
            pictureBox1.Size = new Size(richTextBox1.Size.Height, richTextBox1.Size.Height);
            pan_Select_Path.Dock = DockStyle.Fill;
            tableLayoutPanel1.Dock = DockStyle.Fill;
            materialCheckBox1.Location = new Point(-5, 55);
            materialLabel4.Location = new Point(20, 60);
            button1.Location = new Point(0, 90);
            button3.Location = new Point(button1.Location.X + button1.Size.Width, 90);


            //Дизайн для копирования
            
            materialLabel3.Location = new Point(-5, 5);
            label1.Location = new Point(0, 25);
            progressBar1.Location = new Point(0, 60);
            progressBar1.Size = new Size(550 - 23, 20);
            button2.Location = new Point(402, 90);
            //this.Size = new Size(pictureBox1.Location.X + pictureBox1.Size.Width + 20 , 250);
        }
        private void Copy_directory_Load(object sender, EventArgs e)
        {
            richTextBox1.Text = Properties.Settings.Default.Copy_to;
            richTextBox2.Text = Properties.Settings.Default.Copy_from;
            if (richTextBox1.Text != "")
            {
                materialCheckBox1.Checked = true;
            }
        }
        void copy()
        {
            try
            {
                Directory.Delete(copy_to + @"\", true);
                Copy_dir(Properties.Settings.Default.Copy_from, copy_to + @"\");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
}
        void Copy_dir(string FromDir, string ToDir)
        {

            Directory.CreateDirectory(ToDir);
                foreach (string s1 in Directory.GetFiles(FromDir))
                {
                    string name = Path.GetFileName(s1);
                    //Отображаем прогресс
                    this.BeginInvoke(new Action<string>(file_name =>
                    {
                        label1.Text = file_name;
                        progressBar1.Value += 1;
                        this.Update();
                    }), name);
                    //////
                    string s2 = ToDir + "\\" + name;
                    File.Copy(s1, s2);
                }
                
            
            foreach (string s in Directory.GetDirectories(FromDir))
            {
                Copy_dir(s, ToDir + "\\" + Path.GetFileName(s));
            }
            
        }
        private int Search_dir(string FromDir,int sum_dir)
        {
            foreach (string s1 in Directory.GetFiles(FromDir))
            {
                sum_dir++;
                //string s2 = ToDir + "\\" + Path.GetFileName(s1);
                //File.Copy(s1, s2);
            }
            foreach (string s in Directory.GetDirectories(FromDir))
            {
                sum_dir = Search_dir(s,sum_dir);
            }
            return sum_dir;

        }

        private void materialFlatButton1_Click(object sender, EventArgs e)
        {
          
        }

        private void materialFlatButton2_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (progressBar1.Value == progressBar1.Maximum)
            {
                materialLabel3.Text = "Копирование успешно завершено";
                button2.Enabled = true;
                label1.Text = "";
                this.Update();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
        }
        private void update_pan()
        {
            tableLayoutPanel1.RowStyles[0].Height = 0;
            tableLayoutPanel1.RowStyles[1].Height = 100;
            this.Update();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.Text = fbd.SelectedPath;
                Properties.Settings.Default.Copy_to = fbd.SelectedPath;
                Properties.Settings.Default.Save();
            }    
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = @"C:\";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                richTextBox1.Text = dialog.FileName;
            }
            //OpenFileDialog op = new OpenFileDialog();
            //op.ShowDialog();
        }

        

        private void button1_Click_1(object sender, EventArgs e)
        {
           
            if (richTextBox1.Text != "")
            {
                copy_to = @richTextBox1.Text;
                if (materialCheckBox1.Checked)
                {
                    Properties.Settings.Default.Copy_to = richTextBox1.Text;
                    Properties.Settings.Default.Save();
                }
                else
                {
                    Properties.Settings.Default.Copy_to = "";
                    Properties.Settings.Default.Save();
                }
                update_pan();
                try
                {
                    timer1.Start();
                    int _sum_dir = 0;
                    _sum_dir = Search_dir(Properties.Settings.Default.Copy_from, _sum_dir);
                    progressBar1.Minimum = 0;
                    progressBar1.Maximum = _sum_dir;
                    Thread t = new Thread(copy);
                    t.Start();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Пожалуйста выберете путь для копирования и повторите попытку");
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
