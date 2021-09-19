using System;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToSQL
{
    class Database_query
    {
        private static string SQL_string = @Properties.Settings.Default.SQL_string;
        private SqlConnection sqlcon = null;
        private DataTable dt_save = null;

        //Выполняет запрос на вставку данных созданный функцией Generete_insert_query
        public void Insert_query(DataTable dt, string Table_Name,string mes = "")
        {
            dt_save = BD_load(Table_Name);
            Table_Name = "[" + Properties.Settings.Default.DB_name + "].[dbo].[" + Table_Name + "$]";
            //MessageBox.Show(Generate_insert_query(dt.Rows[0],"test"));
            try
            {
                sqlcon = new SqlConnection(SQL_string);//создание подключения
                sqlcon.Open();
                SqlCommand sql_com_del = new SqlCommand(@"DELETE FROM " + Table_Name, sqlcon);
                try
                {
                    sql_com_del.ExecuteNonQuery();
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Ошибка при удалении данных\nТекст ошибки : " + ex.ToString(), "ОЙ-ОЙ");
                }
                //throw new Exception("ОШИБКААААА:)))");
                foreach (DataRow dr in dt.Rows)
                {
                    SqlCommand sql_com = new SqlCommand(@Generate_insert_query(dr, Table_Name), sqlcon);
                    sql_com.ExecuteNonQuery();
                }
                Mes_from mf = new Mes_from(mes + "OK");
                mf.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при заполнении базы данных\nВ базу данных будет отправлена версия таблицы до изменения\nТекст ошибки : "+ex.ToString(),"ОЙ-ОЙ");
                Recovery_db(Table_Name);
            }
            finally
            {
                sqlcon.Close();

            }

        }
        public string Generate_insert_query(DataRow dr, string Table_Name)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("INSERT INTO ");
            sb.Append(Table_Name);
            sb.Append(" VALUES (");
            foreach (var x in dr.ItemArray)
            {
                if ((x.ToString() != "") && (x.ToString() != " "))
                {
                    sb.Append("'");
                    sb.Append(x.ToString());
                    sb.Append("',");
                }
                else
                {
                    sb.Append("NULL,");
                }
            }
            sb.Remove(sb.Length - 1, 1);
            sb.Append(")");
            return sb.ToString();
        }
        private void Recovery_db(string Table_Name)
        {
            try
            {
                sqlcon = new SqlConnection(SQL_string);//создание подключения
                sqlcon.Open();
                SqlCommand sql_com_del = new SqlCommand(@"DELETE FROM " + Table_Name, sqlcon);
                try
                {
                    sql_com_del.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при удалении данных в процедуре восстановления\nТекст ошибки : " + ex.ToString(), "ОЙ-ОЙ");
                }
                foreach (DataRow dr in dt_save.Rows)
                {
                    SqlCommand sql_com = new SqlCommand(@Generate_insert_query(dr, Table_Name), sqlcon);
                    sql_com.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Попытка заполнения базы данных закончилась неудачей\nТекст ошибки : " + ex.ToString(), "ОЙ-ЁЁЁЁЁЙ");
            }
            finally
            {
                sqlcon.Close();

            }

        }
        //Используется в Path_load
        public static List<string> Tab_name()
        {
            List<string> list_name = new List<string> { };

            try
            {
                DataTable dt = new DataTable();
            SqlConnection sqlcon = new SqlConnection(SQL_string);


                sqlcon.Open();
                SqlDataAdapter sql_adap = new SqlDataAdapter(@"SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES", sqlcon);
                sql_adap.Fill(dt);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    list_name.Add(dt.Rows[i][0].ToString());
                }
            }
            catch (Exception ex)
            {
                
                MessageBox.Show("Ошибка при выоборке названий таблиц.\nПожалуйста проверьте правильность настроек и перезапустите приложение\n Ошибка : " + ex.Message);
                Properties.Settings_edit se = new Properties.Settings_edit("ERROR");
                se.ShowDialog();

            }
            return list_name;
        }
        //в программе не используется из-за смены консепции 
        public static string Person_load(string log, string pass)
        {
            DataTable dt = new DataTable();
            SqlConnection sqlcon = new SqlConnection(SQL_string);
            try
            {
                sqlcon.Open();
                SqlDataAdapter sql_adap = new SqlDataAdapter(@"SELECT * FROM login", sqlcon);
                sql_adap.Fill(dt);
                if (dt.Rows.Count != 0)
                {
                    return dt.Rows[0][0].ToString();
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при попытки проверки логина\n" + ex.Message);
                return null;
            }
        }
        public static bool Update_on(string id, string log, string pass, string table = null)
        {
            DataTable dt = new DataTable();
            SqlConnection sqlcon = new SqlConnection(SQL_string);
            try
            {
                if (table != null)
                {
                    sqlcon.Open();
                    SqlCommand sql_com_del = new SqlCommand(@"UPDATE login SET onlines = 'TRUE', online_table = '" + table + "' WHERE id = " + id + " ;");
                    sql_com_del.ExecuteNonQuery();
                }
                else
                {
                    sqlcon.Open();
                    SqlCommand sql_com_del = new SqlCommand("UPDATE login SET onlines = 1 WHERE id = " + id + " ;", sqlcon);
                    sql_com_del.ExecuteNonQuery();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
        public DataTable BD_load(string TableName)
        {
            try
            {
                DataTable dt = new DataTable();
                SqlConnection sqlcon = new SqlConnection(SQL_string);
                StringBuilder sb = new StringBuilder();
                sb.Append("Select * From [" + Properties.Settings.Default.DB_name + "].[dbo].[");
                sb.Append(TableName);
                sb.Append("$] ORDER BY ISNULL(Item,'zzzzz')");
                sb.Append(";");
                SqlDataAdapter sqladap = new SqlDataAdapter(sb.ToString(), sqlcon);
                sqladap.Fill(dt);
                return dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка во время загрузки данных таблицы : " + TableName + "\nТекст ошибки : \n" + ex.Message,":(");
                return null;
            }
        }

        //создает таблицу по листу если ее не существует
        public static void Create_table(string name, List<string> columns)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("Create table ["+ Properties.Settings.Default.DB_name +"].[dbo].[");
            sb.Append(name);
            sb.Append("$");
            sb.Append("] \n(");
            foreach(var a in columns)
            {
                sb.Append("[");
                sb.Append(a);
                sb.Append("] nvarchar(255),");
            }
            sb.Remove(sb.Length - 1, 1);
            sb.Append(")");
            //Запрос к бд
            SqlConnection sqlcon_stat = new SqlConnection(SQL_string);//создание подключения
            try
            {
                sqlcon_stat.Open();
                //MessageBox.Show(sb.ToString());
                SqlCommand sql_com = new SqlCommand(sb.ToString(),sqlcon_stat);
                sql_com.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при отправке запроса на создание таблицы : "+ name +"\nОшибка : \n"+ ex.ToString());
            }
            finally
            {
                sqlcon_stat.Close();

            }
        }
        public static bool Check_con_string()
        {
            try
            {
                SqlConnection sqlcon = new SqlConnection(SQL_string);
                sqlcon.Open();
                SqlDataAdapter sql_adap = new SqlDataAdapter(@"SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES", sqlcon);
                return true;
            }
            catch 
            {
                return false;
            }
            
        }
        public void RE_Creating_table(DataTable dt, string name)
        {
            List<string> name_list = new List<string>(); 
            foreach(DataColumn f in dt.Columns)
            {
                name_list.Add(f.ColumnName);
            }
            if (delete_table(name + @"$"))
            {
                Create_table(name, name_list);
                Insert_query(dt, name, "CREATE_");
            }
        }

        public static bool delete_table(string table)
        {
            try
            {
                SqlConnection sqlcon = new SqlConnection(SQL_string);
                sqlcon.Open();
                SqlCommand sql_com = new SqlCommand(@"DROP TABLE [" + Properties.Settings.Default.DB_name + "].[dbo].[" + table + "] ;", sqlcon);
                sql_com.ExecuteNonQuery();
                sqlcon.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }
            
    }
    
}
