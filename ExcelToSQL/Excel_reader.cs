using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ExcelToSQL
{
    class Excel_reader
    {
        // Хранит в себе все пути в виде path_dic(Key) = (path, Bd_Name)
        private Dictionary<string, List<string>> path_dic = new Dictionary<string, List<string>> { };

        public Excel_reader()
        {
            Path_load pl = new Path_load();
            path_dic = Path_load.Path_dic;
        }
        //возвращает сформированный datatable по названию листа ехель
        public DataTable GenerateDataTable(string NameSheet)
        {

            int iter = 0;
            try
            {
                ExcelPackage package = new ExcelPackage(new FileInfo(path_dic[NameSheet][0]));
                DataTable dt = new DataTable();
                var sheet = package.Workbook.Worksheets[NameSheet];
                int GLUR = sheet.GetValuedDimension().End.Row;
                // MessageBox.Show(sheet.Dimension.End.Column.ToString());
                for (int rows_num = 1; rows_num < GLUR; rows_num++)
                {
                    dt.Rows.Add();
                    iter++;
                }
                for (int column_num = 1; column_num < sheet.GetValuedDimension().End.Column + 1; column_num++)
                {
                    iter++;
                    if (sheet.Cells[1, column_num].Value != null)
                    {
                        dt.Columns.Add(sheet.Cells[1, column_num].Value.ToString());
                    }
                }
                for (int column_num = 1; column_num < dt.Columns.Count + 1; column_num++)
                {
                    for (int rows_num = 2; rows_num <  GLUR+ 1; rows_num++)
                    {
                        iter++;
                        if (sheet.Cells[rows_num, column_num].Value != null)
                        {
                            dt.Rows[rows_num - 2][column_num - 1] = sheet.Cells[rows_num, column_num].Value.ToString();
                        }
                        else
                        {
                            dt.Rows[rows_num - 2][column_num - 1] = "";
                        }
                    }
                }


                return dt;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Ошибка при попытке вывести лист : " + NameSheet + "\nКод ошибки : \n" + ex.Message, ":o");
                return null;
            }
            
        }

        //функция для заполеннения екселя (пережиток прошлого программы), но если консепция опять поменяется можно использовать
        public void FillExcel(DataTable dt, string Name_sheet)
        {
            ExcelPackage package = new ExcelPackage(new FileInfo(path_dic[Name_sheet][1]));
            ClearExcel(Name_sheet, package);
            var sheet = package.Workbook.Worksheets[Name_sheet];
            sheet.Cells[2, 1].LoadFromDataTable(dt);
            package.Save();
            //Velnus.Message_form mf = new Velnus.Message_form("nice");
            //mf.ShowDialog();

        }
        // Если Hat = true тогда очисткаа происходит без первой строки листа Excel
        // В программе не используется из-за смены консепции
        private void ClearExcel(string Name_sheet, ExcelPackage package, bool Hat = true)
        {
            List<string> ls = new List<string> { };
            int j;
            if (Hat) { j = 2; }
            else { j = 1; }

            var sheet = package.Workbook.Worksheets[Name_sheet];

            for (int column_num = 1; column_num < sheet.Dimension.End.Column; column_num++)
                for (int row_num = j; row_num < sheet.Dimension.End.Row; row_num++)
                {
                    sheet.Cells[row_num, column_num].LoadFromText("");
                }
        }
    }
}
