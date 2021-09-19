using System;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Drawing;

namespace ExcelToSQL
{
    //класс вызывается в excel_reader
    //после можно обращаться к нему по необходимости
    class Path_load
    {

        // Хранит в себе все пути в виде path_dic(Key) = (path, Bd_Name)
        private static Dictionary<string, List<string>> path_dic = new Dictionary<string, List<string>> { };
        public static Dictionary<string, List<string>> no_name_table = new Dictionary<string, List<string>> { };
        //Заполнение пути к файлу из settings.settings
        public static System.Collections.Specialized.StringCollection path_ls = @Properties.Settings.Default.Path_to_folder_library;
        private static bool first_try = true;
        // свойство класса, как никак ООП
        public static Dictionary<string, List<string>> Path_dic { get { return path_dic; } set { } }

        //Полное заполенение словаря с использованием загрузки из бд и сопоставления с названиями слоев из пути 
        public Path_load()
        {
                path_dic.Clear();
                no_name_table.Clear();
                List<string> table_names = Database_query.Tab_name();
                List<string> filesDir1 = new List<string> { };
                foreach (var path in path_ls)
                {
                    filesDir1.AddRange(from a in Directory.GetFiles(path)
                                       where ((!a.Contains("~")) && (a.Contains(".xlsx")))
                                       select a);
                }
                try
                {
                    foreach (string c in filesDir1)
                    {
                        ExcelPackage package = new ExcelPackage(new FileInfo(@c));
                        foreach (var a in package.Workbook.Worksheets)
                        {
                            if (table_names.Contains(a.Name + @"$"))
                            {
                                path_dic.Add(a.Name, new List<string> { c, table_names.First(s => s == a.Name + @"$") });
                            }
                            else
                            {
                                no_name_table.Add(a.Name, new List<string>());
                                for (int i = 1; i < a.GetValuedDimension().End.Column + 1; i++)
                                {
                                    no_name_table[a.Name].Add(a.Cells[1, i].Value.ToString());
                                }
                                //if (!_no_table_item.Contains(a.Name))
                                //{
                                //    result = MessageBox.Show("Вы хотите создать новую таблицу (MSSQL) для листа " + a.Name + "?" + "\nКол-во столбцов : " + a.GetValuedDimension().End.Column
                                //        ,
                                //                "В базе данных не существует таблицы для " + a.Name,
                                //                 MessageBoxButtons.YesNo,
                                //                 MessageBoxIcon.Question);
                                //    if (result == DialogResult.Yes)
                                //    {
                                //        List<string> column_list = new List<string> { };
                                //        for (int i = 1; i < a.GetValuedDimension().End.Column + 1; i++)
                                //        {
                                //            column_list.Add(a.Cells[1, i].Value.ToString());
                                //        }
                                //        Database_query.Create_table(a.Name, column_list);
                                //    }
                                //}
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Произошла ошибка в добавлении листов \n" + ex.Message);
                }
                first_try = false;
        }
        public static Dictionary<string, List<string>> tab_update()
        {
            path_dic.Clear();
            no_name_table.Clear();
            List<string> table_names = Database_query.Tab_name();
            List<string> filesDir1 = new List<string> { };
            foreach (string path in Path_load.path_ls)
            {
                filesDir1.AddRange(from a in Directory.GetFiles(path)
                                   where ((!a.Contains("~")) && (a.Contains(".xlsx")))
                                   select a);
            }
            try
            {
                foreach (string c in filesDir1)
                {
                    ExcelPackage package = new ExcelPackage(new FileInfo(@c));
                    foreach (var a in package.Workbook.Worksheets)
                    {
                        if (table_names.Contains(a.Name + @"$"))
                        {
                        }
                        else
                        {
                            no_name_table.Add(a.Name, new List<string>());
                            for (int i = 1; i < a.GetValuedDimension().End.Column + 1; i++)
                            {
                                no_name_table[a.Name].Add(a.Cells[1, i].Value.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка в добавлении листов \n" + ex.Message);
            }
            return no_name_table;
        }
        //процедура обновления путей к листам
        public void Path_update()
        {
            path_dic.Clear();
            no_name_table.Clear();
            List<string> table_names = Database_query.Tab_name();
            List<string> filesDir1 = new List<string> { };
            foreach (var path in path_ls)
            {
                filesDir1.AddRange(from a in Directory.GetFiles(path)
                                   where ((!a.Contains("~")) && (a.Contains(".xlsx")))
                                   select a);
            }
            foreach (string c in filesDir1)
            {
                ExcelPackage package = new ExcelPackage(new FileInfo(c));
                foreach (var a in package.Workbook.Worksheets)
                {
                    if (table_names.Contains(a.Name + @"$"))
                    {
                        path_dic.Add(a.Name, new List<string> { c, table_names.First(s => s == a.Name + @"$") });
                    }   
                }
            }
        }
    }
}
