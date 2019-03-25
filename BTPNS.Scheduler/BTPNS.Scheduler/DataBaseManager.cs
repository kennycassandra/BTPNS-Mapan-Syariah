using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BTPNS.Scheduler
{
    public class DataBaseManager
    {
        //public string sqlConnection = "Data Source=SP2K13;Initial Catalog=POC;User ID = sa; Password=pass@word1";
        public string sqlConnection = "";
        public SqlCommand cmd;
        public SqlDataReader dReader;
        public SqlTransaction trans;
        public string GetConnString()
        {
            try
            {
                return ConfigurationManager.ConnectionStrings["cnstr"].ToString();
            }
            catch 
            {
                return "Data Source=.;Initial Catalog=IFMS;Integrated Security=True;User ID = sa; Password=pass@word1";
            }
        }
        public void OpenConnection(ref SqlConnection connection, bool IsTrans = false)
        {
            if (connection == null)
            {
                connection = new SqlConnection(GetConnString());
                connection.Open();
                cmd = connection.CreateCommand();
                cmd.CommandTimeout = 0;
                if (IsTrans)
                {
                    trans = connection.BeginTransaction();
                    cmd.Transaction = trans;
                }
            }
            else
            {
                if (connection.State == ConnectionState.Closed)
                {
                    connection = new SqlConnection(GetConnString());
                    connection.Open();
                    cmd = connection.CreateCommand();
                    cmd.CommandTimeout = 0;
                    if (IsTrans)
                    {
                        trans = connection.BeginTransaction();
                        cmd.Transaction = trans;
                    }
                }
            }
        }
        public void CloseConnection(ref SqlConnection connection, bool IsTrans = false)
        {
            if (connection.State == ConnectionState.Open)
            {
                if (IsTrans)
                {
                    trans.Commit();
                }
                connection.Close();
            }
        }

        public void AddInParameter(SqlCommand command, string name, object value)
        {
            SqlParameter parameter = new SqlParameter();
            parameter.ParameterName = name;
            parameter.Value = value;
            parameter.Direction = ParameterDirection.Input;

            command.Parameters.Add(parameter);
        }
        public void AddOutParameter(SqlCommand command, string name, SqlDbType type)
        {
            SqlParameter parameter = new SqlParameter();
            parameter.ParameterName = name;
            parameter.SqlDbType = type;
            parameter.Direction = ParameterDirection.Output;

            command.Parameters.Add(parameter);
        }
        public void CloseDataReader(SqlDataReader dataReader)
        {
            if (dataReader == null)
                return;
            dataReader.Close();
            dataReader.Dispose();
        }

        public DataTable GetDataByProcedure(string SPName, string ParamName, string ParamValue)
        {
            DataTable dt = new DataTable();
            SqlDataReader dr = null;
            cmd.CommandText = SPName;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Clear();
            cmd.Parameters.AddWithValue(ParamName, ParamValue);
            dr = cmd.ExecuteReader();
            dt.Load(dr);
            CloseDataReader(dr);
            return dt;
        }
        public bool isRecordExists(string SPName, string ParamName, string ParamValue)
        {
            SqlDataReader dr = null;
            cmd.CommandText = SPName;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Clear();
            cmd.Parameters.AddWithValue(ParamName, ParamValue);
            dr = cmd.ExecuteReader();
            if (dr.HasRows)
            {
                CloseDataReader(dr);
                return true;
            }
            else
            {
                CloseDataReader(dr);
                return false;
            }
        }
        public bool isRecordExists(string query)
        {
            SqlDataReader dr = null;
            cmd.CommandText = query;
            cmd.CommandType = CommandType.Text;
            dr = cmd.ExecuteReader();
            if (dr.HasRows)
            {
                CloseDataReader(dr);
                return true;
            }
            else
            {
                CloseDataReader(dr);
                return false;
            }
        }

        public string Autocounter(string fieldName, string TableName, string fieldCriteria, string valueCriteria, int LengthOfString, string fieldNameConverted = "")
        {
            SqlDataReader dReader;
            string str = "";
            string str2 = "1";

            cmd.CommandText = "select top 1 " + fieldName + " from " + TableName +
                                   " where " + fieldCriteria + " like '%" + valueCriteria + "%'";
            if (fieldNameConverted != "")
                cmd.CommandText += " order by " + fieldNameConverted + " desc";
            else
                cmd.CommandText += " order by " + fieldName + " desc";

            cmd.CommandType = CommandType.Text;

            dReader = cmd.ExecuteReader();

            if (dReader.HasRows)
            {
                dReader.Read();
                str2 = dReader[fieldName].ToString();

                int StartFrom = valueCriteria.Length;
                int EndUntil = 0;
                EndUntil = str2.Length - StartFrom;

                str2 = (int.Parse(str2.Substring(StartFrom, EndUntil)) + 1).ToString();
            }
            int i = 0;
            while (i < LengthOfString - str2.Length)
            {
                str += "0";
                i += 1;
            }
            str = valueCriteria + str + str2;
            CloseDataReader(dReader);
            return str;

        }

        public string GetValueFromQuery(string Query, string Field_Name)
        {
            string result = "";
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = Query;
            dReader = cmd.ExecuteReader();
            while (dReader.Read())
            {
                result = dReader[Field_Name].ToString();
            }
            CloseDataReader(dReader);
            return result;
        }
    }

}
