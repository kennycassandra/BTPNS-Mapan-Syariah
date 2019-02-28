using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BTPNS.Scheduler
{
    public class RDLCHelper
    {
         public string GeneratePDF(string RDLC, string OutputFileName)
         {  
            // select appropriate contenttype, while binary transfer it identifies filetype  
             string contentType = string.Empty;
            contentType = "application/pdf";
             //if (ddlFileFormat.SelectedValue.Equals(".pdf"))  
             //    contentType = "application/pdf";  
             //if (ddlFileFormat.SelectedValue.Equals(".doc"))  
             //    contentType = "application/ms-word";  
             //if (ddlFileFormat.SelectedValue.Equals(".xls"))  
             //    contentType = "application/xls";  
   
             DataTable dsData = new DataTable();  
             dsData = getReportData();  
   
             string FileName = OutputFileName;  
             string extension;  
             string encoding;  
             string mimeType;  
             string[] streams;  
             Warning[] warnings;  
   
             LocalReport report = new LocalReport();  
             report.ReportPath = Path.Combine(Environment.CurrentDirectory, "RDLC") + "/" + RDLC;
            ReportDataSource rds = new ReportDataSource();  
             rds.Name = "DataSet1";//This refers to the dataset name in the RDLC file  
             rds.Value = dsData;  
             report.DataSources.Add(rds);  
   
             Byte[] mybytes = report.Render("PDF", null,
                    out extension, out encoding,
                    out mimeType, out streams, out warnings); //for exporting to PDF  
            string file_output_url = Path.Combine(Environment.CurrentDirectory, "Output") + FileName;

             using (FileStream fs = File.Create(file_output_url))  
             {  
                 fs.Write(mybytes, 0, mybytes.Length);  
             }

            return file_output_url;
        }  

    }
}
