﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Data;

namespace Printer.Info
{
    public enum PrinterCategory { MClass,IClass,HClass,AClass,EClass, Performance}
    public class PrinterInfo
    {
        public string CID { get; set; }
        public string PackageLabel { get; set; }
        public string AgencyLabel { get; set; }
        public string ModelName { get; set; }
        public string SNRule { get; set; }
        public string Version { get; set; }
    }

    public class InfoManipulator
    {
        private SqlConnection con;
        private string GetTableName(PrinterCategory category)
        {
            switch(category)
            {
                case PrinterCategory.MClass:
                    return "MClassLabelInfo";
                case PrinterCategory.IClass:
                    return "IClassLabelInfo";
                case PrinterCategory.HClass:
                    return "HClassLabelInfo";
                case PrinterCategory.AClass:
                    return "AClassLabelInfo";
                case PrinterCategory.EClass:
                    return "EClassLabelInfo";
                case PrinterCategory.Performance:
                    return "PerformanceLabelInfo";
                default:
                    return string.Empty;
            }
        }

        public InfoManipulator()
        {
            con = new SqlConnection(@"Data Source = ch71w0120; Initial Catalog = PrinterInfo; User ID = PrinterInfo; Password = 123456");
        }

        public int Open()
        {
            try
            {
                con.Open();
                return 0;
            }
            catch(SqlException e)
            {
                return e.ErrorCode;
            }
        }
        public int Close()
        {
            try
            {
                con.Close();
                return 0;
            }
            catch(SqlException e)
            {
                return e.ErrorCode;
            }
        }
        public PrinterInfo GetPrinterInfo(string CID, PrinterCategory category)
        {
            string table_name = GetTableName(category);
            SqlCommand command = new SqlCommand($"SELECT Package,Agency,ModelName,SNRule,Version FROM {table_name} WHERE CID='{CID}'", con);
            if (con.State == System.Data.ConnectionState.Closed)
            {
                con.Open();
            }
            SqlDataReader reader = command.ExecuteReader();
            PrinterInfo info = new PrinterInfo();
            while(reader.Read())
            {
                info.ModelName = reader[2] == DBNull.Value ? string.Empty : (string)reader[2];
                info.AgencyLabel = (string)reader[1];
                info.PackageLabel = (string)reader[0];
                info.CID = CID;
                info.SNRule = reader[3] == DBNull.Value ? string.Empty : (string)reader[3];
                info.Version = reader[4] == DBNull.Value ? string.Empty : (string)reader[4];
            }
            reader.Close();
            return info;
        }
        public void InsertPrinterInfo(PrinterInfo info,PrinterCategory category)
        {
            string table_name = GetTableName(category);
            SqlCommand command = new SqlCommand($"INSERT INTO {table_name} (CID,Agency,Package,ModelName,SNRule,Version) VALUES('{info.CID}','{info.AgencyLabel}','{info.PackageLabel}','{info.ModelName}','{info.SNRule}','{info.Version}')",con);
            if(con.State == System.Data.ConnectionState.Closed)
            {
                con.Open();
            }
            command.ExecuteNonQuery();
        }
        public void UpdatePrintInfo(PrinterInfo info,PrinterCategory category)
        {
            string table_name = GetTableName(category);
            SqlCommand command = new SqlCommand($"UPDATE {table_name} SET Agency='{info.AgencyLabel}',Package='{info.PackageLabel}',ModelName='{info.ModelName}',SNRule='{info.SNRule}',Version='{info.Version}' WHERE CID='{info.CID}'", con);
            if (con.State == System.Data.ConnectionState.Closed)
            {
                con.Open();
            }
            command.ExecuteNonQuery();
        }
        public void DeleteAllInfo(PrinterCategory category)
        {
            string table_name = GetTableName(category);
            SqlCommand command = new SqlCommand($"DELETE FROM {table_name}", con);
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            command.ExecuteNonQuery();
        }
        public DataSet GetAllInfo(PrinterCategory category)
        {
            string tabel_name = GetTableName(category);
            SqlCommand command = new SqlCommand($"SELECT * FROM {tabel_name}", con);
            if( con.State==ConnectionState.Closed)
            {
                con.Open();
            }
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet(tabel_name);
            da.Fill(ds);
            return ds;
        }
        public int GetSNTail(PrinterCategory category)
        {
            string name = Enum.GetName(typeof(PrinterCategory), category);
            SqlCommand command = new SqlCommand($"SELECT {name} FROM SNTail", con);
            if (con.State == System.Data.ConnectionState.Closed)
            {
                con.Open();
            }
            SqlDataReader reader = command.ExecuteReader();
            PrinterInfo info = new PrinterInfo();
            int sn = 0;
            if(reader.Read())
                sn  = (int)reader[0];
            reader.Close();
            SqlCommand command2 = new SqlCommand($"UPDATE SNTail SET {name}={sn+1} WHERE 1=1", con);
            if (con.State == System.Data.ConnectionState.Closed)
            {
                con.Open();
            }
            command2.ExecuteNonQuery();
            return sn;
        }
        public int GetPitneySNTail()
        {
            string name = "Pitney";
            SqlCommand command = new SqlCommand($"SELECT {name} FROM SNTail", con);
            if (con.State == System.Data.ConnectionState.Closed)
            {
                con.Open();
            }
            SqlDataReader reader = command.ExecuteReader();
            PrinterInfo info = new PrinterInfo();
            int sn = 0;
            if (reader.Read())
                sn = (int)reader[0];
            reader.Close();
            SqlCommand command2 = new SqlCommand($"UPDATE SNTail SET {name}={sn + 1} WHERE 1=1", con);
            if (con.State == System.Data.ConnectionState.Closed)
            {
                con.Open();
            }
            command2.ExecuteNonQuery();
            return sn;
        }
        public void WritePitneySNTail(int SNTail)
        {
            string name = "Pitney";
            SqlCommand command = new SqlCommand($"UPDATE SNTail SET {name}={SNTail} WHERE 1=1", con);
            if (con.State == System.Data.ConnectionState.Closed)
            {
                con.Open();
            }
            command.ExecuteNonQuery();
        }
        public int GetWeek()
        {
            SqlCommand command = new SqlCommand($"SELECT Week FROM SNTail", con);
            if (con.State == System.Data.ConnectionState.Closed)
            {
                con.Open();
            }
            SqlDataReader reader = command.ExecuteReader();
            int week = 0;
            if(reader.Read())
                week = (int)reader[0];
            reader.Close();
            return week;
        }
        public void SetWeek(int week)
        {
            SqlCommand command = new SqlCommand($"UPDATE SNTail SET Week={week},MClass=1,EClass=1,IClass=1,HClass=1,AClass=1 WHERE 1=1", con);
            if (con.State == System.Data.ConnectionState.Closed)
            {
                con.Open();
            }
            command.ExecuteNonQuery();
        }
        public void WriteSNTail(int SNTail,PrinterCategory category)
        {
            string name = Enum.GetName(typeof(PrinterCategory), category);
            SqlCommand command = new SqlCommand($"UPDATE SNTail SET {name}={SNTail} WHERE 1=1",con);
            if (con.State == System.Data.ConnectionState.Closed)
            {
                con.Open();
            }
            command.ExecuteNonQuery();
        }
        public void WriteFlag(int flag,PrinterCategory category)
        {
            string name = Enum.GetName(typeof(PrinterCategory), category) + "Flag";
            SqlCommand command = new SqlCommand($"UPDATE SNTail SET {name}={flag} WHERE 1=1", con);
            if (con.State == System.Data.ConnectionState.Closed)
            {
                con.Open();
            }
            command.ExecuteNonQuery();
        }
        public int ReadFlag(PrinterCategory category)
        {
            string name = Enum.GetName(typeof(PrinterCategory), category) + "Flag";
            SqlCommand command = new SqlCommand($"SELECT {name} FROM SNTail", con);
            if (con.State == System.Data.ConnectionState.Closed)
            {
                con.Open();
            }
            SqlDataReader reader = command.ExecuteReader();
            PrinterInfo info = new PrinterInfo();
            int flag = 0;
            if (reader.Read())
                flag = (int)reader[0];
            reader.Close();
            return flag;
        }
    }
}
