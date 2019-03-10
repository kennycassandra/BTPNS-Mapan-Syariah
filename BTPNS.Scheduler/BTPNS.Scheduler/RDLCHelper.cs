using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BTPNS.Scheduler
{
    public class RDLCHelper
    {
        DataBaseManager db = new DataBaseManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlDataReader reader = null;
        Utility util = new Utility();

        public void GeneratePDF()
        {
            try
            {
                DataTable dt = new DataTable();
                db.OpenConnection(ref sqlConn);
                db.cmd.CommandText = "usp_GeneratePDF_ListAkad";
                db.cmd.CommandType = CommandType.StoredProcedure;

                db.cmd.Parameters.Clear();

                reader = db.cmd.ExecuteReader();
                dt.Load(reader);

                db.CloseConnection(ref sqlConn);

                db.OpenConnection(ref sqlConn);
                foreach(DataRow row in dt.Rows)
                {
                    string NomorAkad = util.GetStringValue(row, "NomorAkad");
                    string output1 = GenerateAP3RPDF("AP3R.rdlc", "AP3R_M-Prospera_No_APPID " + NomorAkad, NomorAkad);
                    string output2 = GeneratePersetujuanPembiayaan("PersetujuanPembiayaan.rdlc", NomorAkad, "PP_M-Prospera_No_APPID" + NomorAkad);

                    Console.WriteLine("Nomor Akad : {0}", NomorAkad);
                    Console.WriteLine("Output1 : {0}", output1);
                    Console.WriteLine("Output2 : {0}", output2);

                    db.cmd.CommandText = "usp_Log_GeneratePDF_Insert";
                    db.cmd.CommandType = CommandType.StoredProcedure;
                    db.cmd.Parameters.Clear();

                    db.AddInParameter(db.cmd, "NomorAkad", NomorAkad);
                    db.AddInParameter(db.cmd, "OutputName1", output1);
                    db.AddInParameter(db.cmd, "OutputName2", output2);

                    db.cmd.ExecuteNonQuery();

                }
                db.CloseConnection(ref sqlConn);
                Console.WriteLine("Success");
            }
            catch (Exception ex)
            {
                db.CloseConnection(ref sqlConn);
                Console.WriteLine(ex);
                throw ex;
            }
        }

        public string GenerateSMS()
        {
            try
            {

                DataTable dt = new DataTable();
                string NoRek, NamaNasabah, NoHp, NomorAkad;
                db.OpenConnection(ref sqlConn);
                db.cmd.CommandText = "usp_SMS_Notification_Generate";
                db.cmd.CommandType = CommandType.StoredProcedure;
                db.cmd.Parameters.Clear();
                reader = db.cmd.ExecuteReader();
                dt.Load(reader);
                db.CloseDataReader(reader);
                db.CloseConnection(ref sqlConn);

                string sms_file_name = "MB_IF_" + DateTime.Now.ToString("yyyyMMdd") + dt.Rows.Count.ToString() + ".txt";
                string file_output_url = Path.Combine(Environment.CurrentDirectory, "Output") + "\\SMS\\" + sms_file_name;

                db.OpenConnection(ref sqlConn, true);
                foreach (DataRow row in dt.Rows)
                {
                    string TextSMS = "";
                    NoRek = util.GetStringValue(row, "NomorRekening");
                    NamaNasabah = util.GetStringValue(row, "NamaNasabah");
                    NoHp = util.GetStringValue(row, "NoHp");
                    NomorAkad = util.GetStringValue(row, "NomorAkad");
                    //20181127|MMS|0812345678|Assalaamu'alaikum. Rekening tabungan Ibu KISWANINGSIH sudah diinput, dg no rek 1234567890. Mohon segera informasikan pada nasabah jika pembiayaan sdh disetujui
                    using (StreamWriter writer = new StreamWriter(file_output_url, true))
                    {
                        TextSMS = "Assalaamu'alaikum. Rekening tabungan Bpk/Ibu " + NamaNasabah + " sudah diinput, dgn no rek " + NoRek + ". Mohon segera informasikan pada nasabah jika pembiayaan sdh disetujui";
                        writer.WriteLine(DateTime.Now.ToString("yyyyMMdd") + "|" + "NASABAH" + "|" + NoHp + "|" + TextSMS);
                    }

                    db.cmd.CommandText = "usp_SMS_Notification_Insert";
                    db.cmd.CommandType = CommandType.StoredProcedure;
                    db.cmd.Parameters.Clear();
                    db.AddInParameter(db.cmd, "NomorAkad", NomorAkad);
                    db.AddInParameter(db.cmd, "ScriptSMS", TextSMS);
                    db.AddInParameter(db.cmd, "TxtFile", sms_file_name);
                    db.cmd.ExecuteNonQuery();

                }
                db.CloseConnection(ref sqlConn, true);

                return file_output_url;
            }
            catch (Exception ex)
            {
                db.CloseConnection(ref sqlConn);
                throw ex;
            }
        }

        public string GeneratePersetujuanPembiayaan(string RDLC, string NomorAkad, string OutputFileName)
        {
            try
            {
                DataTable dt = new DataTable();
                db.OpenConnection(ref sqlConn);
                db.cmd.CommandText = "usp_GeneratePDF_PersetujuanPembiayaan";
                db.cmd.CommandType = CommandType.StoredProcedure;
                db.cmd.Parameters.Clear();

                db.AddInParameter(db.cmd, "NomorAkad", NomorAkad);

                reader = db.cmd.ExecuteReader();
                dt.Load(reader);
                db.CloseDataReader(reader);

                db.CloseConnection(ref sqlConn);

                string contentType = string.Empty;
                contentType = "application/pdf";

                string FileName = OutputFileName;
                string extension;
                string encoding;
                string mimeType;
                string[] streams;
                Warning[] warnings;

                LocalReport report = new LocalReport();
                report.ReportPath = Path.Combine(Environment.CurrentDirectory, "RDLC") + "\\" + RDLC;
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "PP_DataSet";//This refers to the dataset name in the RDLC file  
                rds.Value = dt;
                report.DataSources.Add(rds);


                Byte[] mybytes = report.Render("PDF", null,
                        out extension, out encoding,
                        out mimeType, out streams, out warnings); //for exporting to PDF  
                string file_output_url = Path.Combine(Environment.CurrentDirectory, "Output") + "\\PersetujuanPembiayaan\\" + FileName + ".pdf";

                using (FileStream fs = File.Create(file_output_url))
                {
                    fs.Write(mybytes, 0, mybytes.Length);
                }
                return file_output_url;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

         public string GenerateAP3RPDF(string RDLC, string OutputFileName, string NomorAkad)
         {
            try
            {
                DataTable dt = new DataTable();
                DataTable barang1 = new DataTable();
                DataTable barang2 = new DataTable();

                db.OpenConnection(ref sqlConn);
                db.cmd.CommandText = "usp_GeneratePDF_Ap3R";
                db.cmd.CommandType = CommandType.StoredProcedure;
                db.cmd.Parameters.Clear();
                db.AddInParameter(db.cmd, "NomorAkad", NomorAkad);
                reader = db.cmd.ExecuteReader();
                dt.Load(reader);
                db.CloseDataReader(reader);

                db.cmd.CommandText = "usp_GeneratePDF_AP3R_BarangYangDibiayai1";
                db.cmd.CommandType = CommandType.StoredProcedure;
                db.cmd.Parameters.Clear();
                db.AddInParameter(db.cmd, "NomorAkad", NomorAkad);
                reader = db.cmd.ExecuteReader();
                barang1.Load(reader);
                db.CloseDataReader(reader);

                db.cmd.CommandText = "usp_GeneratePDF_AP3R_BarangYangDibiayai2";
                db.cmd.CommandType = CommandType.StoredProcedure;
                db.cmd.Parameters.Clear();
                db.AddInParameter(db.cmd, "NomorAkad", NomorAkad);
                reader = db.cmd.ExecuteReader();
                barang2.Load(reader);
                db.CloseDataReader(reader);


                db.CloseConnection(ref sqlConn);


                string contentType = string.Empty;
                contentType = "application/pdf";

                string FileName = OutputFileName;
                string extension;
                string encoding;
                string mimeType;
                string[] streams;
                Warning[] warnings;

                LocalReport report = new LocalReport();
                report.ReportPath = Path.Combine(Environment.CurrentDirectory, "RDLC") + "\\" + RDLC;
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "APER_DataSet";//This refers to the dataset name in the RDLC file  
                rds.Value = dt;
                report.DataSources.Add(rds);

                rds = new ReportDataSource();
                rds.Name = "Barang1_DataSet";
                rds.Value = barang1;

                report.DataSources.Add(rds);

                rds = new ReportDataSource();
                rds.Name = "Barang2_DataSet";
                rds.Value = barang2;

                report.DataSources.Add(rds);


                Byte[] mybytes = report.Render("PDF", null,
                        out extension, out encoding,
                        out mimeType, out streams, out warnings); //for exporting to PDF  
                string file_output_url = Path.Combine(Environment.CurrentDirectory, "Output") + "\\AP3R\\" + FileName + ".pdf";

                using (FileStream fs = File.Create(file_output_url))
                {
                    fs.Write(mybytes, 0, mybytes.Length);
                }
                return file_output_url;
            }
            catch (Exception ex)
            {
                db.CloseConnection(ref sqlConn);
                throw ex;
            }
 
        }  

    }
}
