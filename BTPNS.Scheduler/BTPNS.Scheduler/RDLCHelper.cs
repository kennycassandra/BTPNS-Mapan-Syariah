using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.Configuration;
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

        public void GeneratePDF(string OutputFolder, string SPSite)
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

                db.CloseDataReader(reader);
                db.CloseConnection(ref sqlConn);

                foreach(DataRow row in dt.Rows)
                {
                    string OfficerMail = util.GetStringValue(row, "OfficerEmail");
                    string NomorAkad = util.GetStringValue(row, "NomorAkad");
                    string NomorDraft = util.GetStringValue(row, "NomorDraft");
                    string output1 = GenerateAP3RPDF(OutputFolder, "AP3R.rdlc", "AP3R_MapanSyariah_" + NomorAkad, NomorAkad);
                    string output2 = GeneratePersetujuanPembiayaan(OutputFolder, "PersetujuanPembiayaan.rdlc", NomorAkad, "PP_MapanSyariah_" + NomorAkad);
                    string CIF = util.GetStringValue(row, "CIF");
                    string Url_AP3R = new SharePointHelper().UploadFileToDocLib(OutputFolder, output1, "AP3R");
                    string Url_PP = new SharePointHelper().UploadFileToDocLib(OutputFolder, output2, "Persetujuan_Pembiayaan");


                    Console.WriteLine("Nomor Akad : {0}", NomorAkad);
                    Console.WriteLine("Output1 : {0}", output1);
                    Console.WriteLine("Output2 : {0}", output2);

                    Console.WriteLine("Url AP3R : {0}", Url_AP3R);
                    Console.WriteLine("Url Persetujuan Pembiayaan : {0}", Url_PP);


                    db.OpenConnection(ref sqlConn);

                    db.cmd.CommandText = "usp_Log_GeneratePDF_Insert";
                    db.cmd.CommandType = CommandType.StoredProcedure;
                    db.cmd.Parameters.Clear();

                    db.AddInParameter(db.cmd, "NomorAkad", NomorAkad);
                    db.AddInParameter(db.cmd, "OutputName1", SPSite + Url_AP3R);
                    db.AddInParameter(db.cmd, "OutputName2", SPSite + Url_PP);
                    db.AddInParameter(db.cmd, "NomorDraft", NomorDraft);

                    db.cmd.ExecuteNonQuery();
                    db.CloseConnection(ref sqlConn);

                    List<string> listAttach = new List<string>();
                    listAttach.Add(output1);
                    listAttach.Add(output2);

                    new MailHelper().email_send(listAttach, 
                        "AP3R & Persetujuan Pembiayaan -" + CIF + "-" + NomorAkad, OfficerMail, 
                        "AP3R - " + SPSite + Url_AP3R + Environment.NewLine + "Persetujuan Pembiayaan - " + SPSite + Url_PP);

                    Console.WriteLine("Generate PDF AP3R & Form Persetujuan Pembiayaan Nomor Akad {0} Done", NomorAkad);

                }
                Console.WriteLine("Generate PDF {0} Rows Done", dt.Rows.Count);

            }
            catch (Exception ex)
            {
                db.CloseConnection(ref sqlConn);
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "GeneratePDF");
            }
        }

        public string GenerateSMS(string OutputFolder)
        {
            try
            {

                DataTable dt = new DataTable();
                string NoRek, NamaNasabah, NoHp, NomorAkad, NomorDraft;
                db.OpenConnection(ref sqlConn);
                db.cmd.CommandText = "usp_SMS_Notification_Generate";
                db.cmd.CommandType = CommandType.StoredProcedure;
                db.cmd.Parameters.Clear();
                reader = db.cmd.ExecuteReader();
                dt.Load(reader);
                db.CloseDataReader(reader);
                db.CloseConnection(ref sqlConn);

                string sms_file_name = "MB_IF_" + DateTime.Now.ToString("yyyyMMdd") + "_" + dt.Rows.Count.ToString() + ".txt";
                string file_output_url = OutputFolder + "Output" + "\\SMS\\" + sms_file_name;

                db.OpenConnection(ref sqlConn, true);
                foreach (DataRow row in dt.Rows)
                {
                    string TextSMS = "";
                    NoRek = util.GetStringValue(row, "NomorRekening");
                    NamaNasabah = util.GetStringValue(row, "NamaNasabah");
                    NoHp = util.GetStringValue(row, "NoHp").Replace("-","");
                    NomorAkad = util.GetStringValue(row, "NomorAkad");
                    NomorDraft = util.GetStringValue(row, "NomorDraft");
                    //20181127|MMS|0812345678|Assalaamu'alaikum. Rekening tabungan Ibu KISWANINGSIH sudah diinput, dg no rek 1234567890. Mohon segera informasikan pada nasabah jika pembiayaan sdh disetujui
                    using (StreamWriter writer = new StreamWriter(file_output_url, true))
                    {
                        TextSMS = "Assalaamu'alaikum. Dokumen Pembiayaan Bpk/Ibu " + NamaNasabah + " telah kami terima dan segera kami proses. Mohon tunggu kabar selanjutnya dari petugas kami. Terima Kasih.";
                        writer.WriteLine(DateTime.Now.ToString("yyyyMMdd") + "|" + "NASABAH" + "|" + NoHp + "|" + TextSMS);
                    }
                    string Url = new SharePointHelper().UploadFileToDocLib(OutputFolder, file_output_url, "SMS_Notification");

                    db.cmd.CommandText = "usp_SMS_Notification_Insert";
                    db.cmd.CommandType = CommandType.StoredProcedure;
                    db.cmd.Parameters.Clear();
                    db.AddInParameter(db.cmd, "NomorAkad", NomorAkad);
                    db.AddInParameter(db.cmd, "ScriptSMS", TextSMS);
                    db.AddInParameter(db.cmd, "TxtFile", sms_file_name);
                    db.AddInParameter(db.cmd, "SharePointUrl", Url);
                    db.AddInParameter(db.cmd, "NomorDraft", NomorDraft);
                    db.cmd.ExecuteNonQuery();

                }
                db.CloseConnection(ref sqlConn, true);

                if (dt.Rows.Count > 0) Console.WriteLine("Generate SMS Done");

                return file_output_url;

            }
            catch (Exception ex)
            {
                db.CloseConnection(ref sqlConn);
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "GenerateSMS");
                return "";
            }
        }

        public string GeneratePersetujuanPembiayaan(string OutputFolder, string RDLC, string NomorAkad, string OutputFileName)
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
                report.ReportPath = OutputFolder + "RDLC" + "\\" + RDLC;
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "PP_DataSet";//This refers to the dataset name in the RDLC file  
                rds.Value = dt;
                report.DataSources.Add(rds);


                Byte[] mybytes = report.Render("PDF", null,
                        out extension, out encoding,
                        out mimeType, out streams, out warnings); //for exporting to PDF  
                string file_output_url = OutputFolder + "Output" + "\\PersetujuanPembiayaan\\" + FileName + ".pdf";

                using (FileStream fs = File.Create(file_output_url))
                {
                    fs.Write(mybytes, 0, mybytes.Length);
                }
                Console.WriteLine("Generate Persetujuan Pembiayaan {0} Done", NomorAkad);
                return file_output_url;

            }
            catch (Exception ex)
            {
                db.CloseConnection(ref sqlConn);
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "GeneratePersetujuanPembiayaan");
                return "";
            }
        }

        public string GenerateAP3RPDF(string OutputFolder, string RDLC, string OutputFileName, string NomorAkad)
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
                report.ReportPath = OutputFolder + "RDLC" + "\\" + RDLC;
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
                string file_output_url = OutputFolder + "\\Output\\AP3R\\" + FileName + ".pdf";

                using (FileStream fs = File.Create(file_output_url))
                {
                    fs.Write(mybytes, 0, mybytes.Length);
                }
                Console.WriteLine("Generate AP3R {0} Done!", NomorAkad);
                return file_output_url;
            }
            catch (Exception ex)
            {
                db.CloseConnection(ref sqlConn);
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "GenerateAP3RPDF");
                return "";
            }
 
        }  

        public void GenerateExcelSummaryReport_Detail1(string OutputFolder)
        {
            try
            {
                DataTable dt = new DataTable();
                db.OpenConnection(ref sqlConn);
                db.cmd.CommandText = "usp_List_Excel_SummaryReport_Detail1";
                db.cmd.CommandType = CommandType.StoredProcedure;
                db.cmd.Parameters.Clear();
                reader = db.cmd.ExecuteReader();
                dt.Load(reader);
                db.CloseDataReader(reader);
                db.CloseConnection(ref sqlConn);
                string RequestId = "";
                string OfficerEmail = "";
                int i = 0;
                DataTable dtRDLC = new DataTable();

                #region Init Column
                dtRDLC.Columns.Add("SysId", typeof(string));
                dtRDLC.Columns.Add("RequestId", typeof(string));
                dtRDLC.Columns.Add("OfficerEmail", typeof(string));
                dtRDLC.Columns.Add("Bulan", typeof(string));
                dtRDLC.Columns.Add("Tahun", typeof(int));
                dtRDLC.Columns.Add("Wisma", typeof(string));
                dtRDLC.Columns.Add("NasabahPotensial", typeof(int));
                dtRDLC.Columns.Add("KonfirmasiNasabah", typeof(int));
                dtRDLC.Columns.Add("StatusPengajuanNasabah", typeof(int));
                dtRDLC.Columns.Add("CarryOver", typeof(int));
                dtRDLC.Columns.Add("EmailSend", typeof(int));
                dtRDLC.Columns.Add("RequestDate", typeof(DateTime));
                #endregion

                foreach (DataRow row in dt.Rows)
                {
                    if (i == 0)
                    {
                        RequestId = util.GetStringValue(row, "RequestId");
                        OfficerEmail = util.GetStringValue(row, "OfficerEmail");
                        dtRDLC.Rows.Add(row.ItemArray);
                    }
                    else
                    {
                        if (RequestId != util.GetStringValue(row, "RequestId"))
                        {
                            RequestId = util.GetStringValue(row, "RequestId");
                            OfficerEmail = util.GetStringValue(row, "OfficerEmail");
                            break;
                        }
                        else
                        {
                            dtRDLC.Rows.Add(row.ItemArray);
                        }
                    }
                    i++;
                }
                string FileName = "SummaryStatusReport_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                string file_output_url = "";
                if (dt.Rows.Count > 0)
                {
                    #region RDLC
                    string extension;
                    string encoding;
                    string mimeType;
                    string[] streams;
                    Warning[] warnings;

                    LocalReport report = new LocalReport();
                    report.ReportPath = OutputFolder + "RDLC" + "\\SummaryStatusMapanSyariah.rdlc";
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = "SummaryStatus_DS";//This refers to the dataset name in the RDLC file  
                    rds.Value = dtRDLC;
                    report.DataSources.Add(rds);

                    Byte[] mybytes = report.Render("Excel", null,
                            out extension, out encoding,
                            out mimeType, out streams, out warnings); //for exporting to PDF  
                    file_output_url = OutputFolder + "Output" + "\\Daily\\" + FileName;

                    using (FileStream fs = File.Create(file_output_url))
                    {
                        fs.Write(mybytes, 0, mybytes.Length);
                    }
                    #endregion
                }

                #region Send Email
                if (dt.Rows.Count > 0)
                {
                    List<string> list = new List<string>();
                    list.Add(file_output_url);
                    new MailHelper().email_send(list, "Summary Status Mapan Syariah Report", OfficerEmail);
                }
                #endregion

                #region Update Flag Send Email
                db.OpenConnection(ref sqlConn);
                db.cmd.CommandText = "Update Excel_SummaryReport_Detail1 set EmailSend=1 Where RequestId = '" + RequestId + "'";
                db.cmd.CommandType = CommandType.Text;
                db.cmd.ExecuteNonQuery();
                db.CloseConnection(ref sqlConn);
                #endregion

                if (dt.Rows.Count > 0)
                {
                    Console.WriteLine("GenerateExcelSummaryReport_Detail1 Done");
                }
            }
            catch (Exception ex)
            {
                db.CloseConnection(ref sqlConn);
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "GenerateExcelSummaryReport_Detail1");
            }
        }

        public void GenerateExcelDetailReport(string OutputFolder)
        {
            try
            {
                DataTable dt = new DataTable();
                db.OpenConnection(ref sqlConn);
                db.cmd.CommandText = "usp_List_Excel_DetailReport";
                db.cmd.CommandType = CommandType.StoredProcedure;
                db.cmd.Parameters.Clear();
                reader = db.cmd.ExecuteReader();
                dt.Load(reader);
                db.CloseDataReader(reader);
                db.CloseConnection(ref sqlConn);
                string RequestId = "";
                string OfficerEmail = "";
                int i = 0;
                DataTable dtRDLC = new DataTable();

                #region Init Column
                dtRDLC.Columns.Add("SysId", typeof(string));
                dtRDLC.Columns.Add("RequestId", typeof(string));
                dtRDLC.Columns.Add("OfficerEmail", typeof(string));
                dtRDLC.Columns.Add("Wisma", typeof(string));
                dtRDLC.Columns.Add("Sentra", typeof(string));
                dtRDLC.Columns.Add("CIF", typeof(string));
                dtRDLC.Columns.Add("NamaNasabah", typeof(string));
                dtRDLC.Columns.Add("Tenor", typeof(int));
                dtRDLC.Columns.Add("Plafon", typeof(decimal));
                dtRDLC.Columns.Add("TglSurvey", typeof(string));
                dtRDLC.Columns.Add("JmlFollowUp", typeof(int));
                dtRDLC.Columns.Add("KonfirmasiNasabah", typeof(string));
                dtRDLC.Columns.Add("TglGenerateCIF", typeof(string));
                dtRDLC.Columns.Add("TglGeneratePembiayaan", typeof(string));
                dtRDLC.Columns.Add("SCO", typeof(string));
                dtRDLC.Columns.Add("StatusPengajuanNasabah", typeof(string));
                dtRDLC.Columns.Add("EmailSend", typeof(int));
                dtRDLC.Columns.Add("RequestDate", typeof(DateTime));
                dtRDLC.Columns.Add("GenerateCIFBy", typeof(string));
                dtRDLC.Columns.Add("GeneratePembiayaanBy", typeof(string));

                #endregion

                foreach (DataRow row in dt.Rows)
                {
                    if (i == 0)
                    {
                        RequestId = util.GetStringValue(row, "RequestId");
                        OfficerEmail = util.GetStringValue(row, "OfficerEmail");
                        dtRDLC.Rows.Add(row.ItemArray);
                    }
                    else
                    {
                        if (RequestId != util.GetStringValue(row, "RequestId"))
                        {
                            RequestId = util.GetStringValue(row, "RequestId");
                            OfficerEmail = util.GetStringValue(row, "OfficerEmail");
                            break;
                        }
                        else
                        {
                            dtRDLC.Rows.Add(row.ItemArray);
                        }
                    }
                    i++;
                }


                string FileName = "DetailStatusMapanSyariahReport_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                string file_output_url = "";
                if (dt.Rows.Count > 0)
                {
                    #region RDLC
                    string extension;
                    string encoding;
                    string mimeType;
                    string[] streams;
                    Warning[] warnings;

                    LocalReport report = new LocalReport();
                    report.ReportPath = OutputFolder + "RDLC" + "\\DetailStatusReportMapanSyariah.rdlc";
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = "DetailReport_DS";
                    rds.Value = dtRDLC;
                    report.DataSources.Add(rds);

                    Byte[] mybytes = report.Render("Excel", null,
                            out extension, out encoding,
                            out mimeType, out streams, out warnings); //for exporting to PDF  
                    file_output_url = OutputFolder + "Output" + "\\Daily\\" + FileName;

                    using (FileStream fs = File.Create(file_output_url))
                    {
                        fs.Write(mybytes, 0, mybytes.Length);
                    }
                    #endregion
                }
                #region Send Email
                if (dt.Rows.Count > 0)
                {
                    List<string> list = new List<string>();
                    list.Add(file_output_url);
                    new MailHelper().email_send(list, "Detail Status Mapan Syariah Report", OfficerEmail);
                }
                #endregion

                #region Update Flag Send Email
                db.OpenConnection(ref sqlConn);
                db.cmd.CommandText = "Update Excel_DetailReport set EmailSend=1 Where RequestId = '" + RequestId + "'";
                db.cmd.CommandType = CommandType.Text;
                db.cmd.ExecuteNonQuery();
                db.CloseConnection(ref sqlConn);
                #endregion

                Console.WriteLine("Generate Excel Detail Report Done");
            }
            catch (Exception ex)
            {
                db.CloseConnection(ref sqlConn);
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "GenerateExcelDetailReport");
            }
        }    

        public void GenerateExcelSummaryReport_Detail2(string OutputFolder)
        {
            try
            {
                DataTable dt = new DataTable();
                db.OpenConnection(ref sqlConn);
                db.cmd.CommandText = "usp_List_Excel_SummaryReport_Detail2";
                db.cmd.CommandType = CommandType.StoredProcedure;
                db.cmd.Parameters.Clear();
                reader = db.cmd.ExecuteReader();
                dt.Load(reader);
                db.CloseDataReader(reader);
                db.CloseConnection(ref sqlConn);
                string RequestId = "";
                string OfficerEmail = "";
                int i = 0;
                DataTable dtRDLC = new DataTable();

                #region Init Column
                dtRDLC.Columns.Add("SysId", typeof(string));
                dtRDLC.Columns.Add("RequestId", typeof(string));
                dtRDLC.Columns.Add("OfficerEmail", typeof(string));
                dtRDLC.Columns.Add("Bulan", typeof(string));
                dtRDLC.Columns.Add("Tahun", typeof(int));
                dtRDLC.Columns.Add("Wism", typeof(string));
                dtRDLC.Columns.Add("BelumDikunjungi", typeof(int));
                dtRDLC.Columns.Add("BertemuNasabah", typeof(int));
                dtRDLC.Columns.Add("TidakBertemu", typeof(int));
                dtRDLC.Columns.Add("DropOff", typeof(int));
                dtRDLC.Columns.Add("Lanjut", typeof(int));
                dtRDLC.Columns.Add("MenungguApproval", typeof(int));
                dtRDLC.Columns.Add("Disetujui", typeof(int));
                dtRDLC.Columns.Add("Ditolak", typeof(int));
                dtRDLC.Columns.Add("MenungguApprovalCO", typeof(int));
                dtRDLC.Columns.Add("DisetujuiCO", typeof(int));
                dtRDLC.Columns.Add("DitolakCO", typeof(int));
                dtRDLC.Columns.Add("EmailSend", typeof(int));
                dtRDLC.Columns.Add("RequestDate", typeof(DateTime));

                #endregion

                foreach (DataRow row in dt.Rows)
                {
                    if (i == 0)
                    {
                        RequestId = util.GetStringValue(row, "RequestId");
                        OfficerEmail = util.GetStringValue(row, "OfficerEmail");
                        dtRDLC.Rows.Add(row.ItemArray);
                    }
                    else
                    {
                        if (RequestId != util.GetStringValue(row, "RequestId"))
                        {
                            RequestId = util.GetStringValue(row, "RequestId");
                            OfficerEmail = util.GetStringValue(row, "OfficerEmail");
                            break;
                        }
                        else
                        {
                            dtRDLC.Rows.Add(row.ItemArray);
                        }
                    }
                    i++;
                }

                string FileName = "DetailSummaryStatusMapanSyariahReport_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                string file_output_url = OutputFolder + "Output" + "\\Daily\\" + FileName;

                if (dt.Rows.Count > 0)
                {
                    #region RDLC
                    string extension;
                    string encoding;
                    string mimeType;
                    string[] streams;
                    Warning[] warnings;

                    LocalReport report = new LocalReport();
                    report.ReportPath = OutputFolder + "RDLC" + "\\DetailSummaryStatusMapanSyariah.rdlc";
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = "DetailStatus_DS";
                    rds.Value = dtRDLC;
                    report.DataSources.Add(rds);

                    Byte[] mybytes = report.Render("Excel", null,
                            out extension, out encoding,
                            out mimeType, out streams, out warnings); //for exporting to PDF  

                    using (FileStream fs = File.Create(file_output_url))
                    {
                        fs.Write(mybytes, 0, mybytes.Length);
                    }
                    #endregion
                }

                #region Send Email
                if (dt.Rows.Count > 0)
                {
                    List<string> list = new List<string>();
                    list.Add(file_output_url);
                    new MailHelper().email_send(list, "Detail Summary Status Mapan Syariah Report", OfficerEmail);
                }
                #endregion

                #region Update Flag Send Email
                db.OpenConnection(ref sqlConn);
                db.cmd.CommandText = "Update Excel_SummaryReport_Detail2 set EmailSend=1 Where RequestId = '" + RequestId + "'";
                db.cmd.CommandType = CommandType.Text;
                db.cmd.ExecuteNonQuery();
                db.CloseConnection(ref sqlConn);
                #endregion
                if (dt.Rows.Count > 0) Console.WriteLine("Generate Excel Summary Report Detail2 Done");

            }
            catch (Exception ex)
            {
                db.CloseConnection(ref sqlConn);
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "GenerateExcelSummaryReport_Detail2");
            }
        }

        public void GenerateExcelLogReport(string OutputFolder)
        {
            try
            {
                DataTable dt = new DataTable();
                db.OpenConnection(ref sqlConn);
                db.cmd.CommandText = "usp_List_Excel_LogManagement";
                db.cmd.CommandType = CommandType.StoredProcedure;
                db.cmd.Parameters.Clear();
                reader = db.cmd.ExecuteReader();
                dt.Load(reader);
                db.CloseDataReader(reader);
                db.CloseConnection(ref sqlConn);
                string RequestId = "";
                string OfficerEmail = "";
                int i = 0;
                DataTable dtRDLC = new DataTable();

                #region Init Column
                dtRDLC.Columns.Add("SysId", typeof(string));
                dtRDLC.Columns.Add("RequestId", typeof(string));
                dtRDLC.Columns.Add("OfficerEmail", typeof(string));
                dtRDLC.Columns.Add("Wisma", typeof(string));
                dtRDLC.Columns.Add("Sentra", typeof(string));
                dtRDLC.Columns.Add("CIF", typeof(string));
                dtRDLC.Columns.Add("Event", typeof(string));
                dtRDLC.Columns.Add("JenisTransaksi", typeof(string));
                dtRDLC.Columns.Add("DiajukanOleh", typeof(string));
                dtRDLC.Columns.Add("TglJamPengajuan", typeof(string));
                dtRDLC.Columns.Add("DiubahOleh", typeof(string));
                dtRDLC.Columns.Add("TglJamPerubahan", typeof(string));
                dtRDLC.Columns.Add("EmailSend", typeof(int));
                dtRDLC.Columns.Add("RequestDate", typeof(DateTime));

                #endregion

                foreach (DataRow row in dt.Rows)
                {
                    if (i == 0)
                    {
                        RequestId = util.GetStringValue(row, "RequestId");
                        OfficerEmail = util.GetStringValue(row, "OfficerEmail");
                        dtRDLC.Rows.Add(row.ItemArray);
                    }
                    else
                    {
                        if (RequestId != util.GetStringValue(row, "RequestId"))
                        {
                            RequestId = util.GetStringValue(row, "RequestId");
                            OfficerEmail = util.GetStringValue(row, "OfficerEmail");
                            break;
                        }
                        else
                        {
                            dtRDLC.Rows.Add(row.ItemArray);
                        }
                    }
                    i++;
                }

                #region RDLC
                string FileName = "";
                string file_output_url = "";
                if (dt.Rows.Count > 0)
                {
                    FileName = "LogReport_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                    string extension;
                    string encoding;
                    string mimeType;
                    string[] streams;
                    Warning[] warnings;

                    LocalReport report = new LocalReport();
                    report.ReportPath = OutputFolder + "RDLC" + "\\LogManagementReport.rdlc";
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = "LogReport_DS";
                    rds.Value = dtRDLC;
                    report.DataSources.Add(rds);

                    Byte[] mybytes = report.Render("Excel", null,
                            out extension, out encoding,
                            out mimeType, out streams, out warnings); //for exporting to PDF  
                    file_output_url = OutputFolder + "Output" + "\\Daily\\" + FileName;

                    using (FileStream fs = File.Create(file_output_url))
                    {
                        fs.Write(mybytes, 0, mybytes.Length);
                    }
                }
                #endregion

                #region Send Email
                if (dt.Rows.Count > 0)
                {
                    List<string> list = new List<string>();
                    list.Add(file_output_url);
                    new MailHelper().email_send(list, "Log Management Report", OfficerEmail);
                }
                #endregion

                #region Update Flag Send Email
                db.OpenConnection(ref sqlConn);
                db.cmd.CommandText = "Update Excel_LogManagement set EmailSend=1 Where RequestId = '" + RequestId + "'";
                db.cmd.CommandType = CommandType.Text;
                db.cmd.ExecuteNonQuery();
                db.CloseConnection(ref sqlConn);
                #endregion

                Console.WriteLine("Generate Excel Log Report Done");
            }
            catch (Exception ex)
            {
                db.CloseConnection(ref sqlConn);
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "GenerateExcelLogReport");
            }

        }

    }
}
