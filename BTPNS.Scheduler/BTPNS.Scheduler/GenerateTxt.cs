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
    public class EmailTxt
    {
        public string Email { get; set; }
        public string NomorDraft { get; set; }
        public string file_attachment { get; set; }
    }
    public class GenerateTxt
    {
        DataBaseManager db = new DataBaseManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlDataReader reader = null;
        Utility util = new Utility();

        public List<EmailTxt> GenerateTxtPembiayaan()
        {
            DataTable dt = new DataTable();
            DataTable dtDetail = new DataTable();
            List<EmailTxt> list = new List<EmailTxt>();
            try
            {
                db.OpenConnection(ref sqlConn);
                db.cmd.CommandText = "usp_List_UnGenerate_CIF_Pembiayaan_PerAgent";
                db.cmd.CommandType = CommandType.StoredProcedure;
                db.cmd.Parameters.Clear();
                db.AddInParameter(db.cmd, "GenerateType", "Pembiayaan");
                reader = db.cmd.ExecuteReader();
                dt.Load(reader);
                db.CloseDataReader(reader);
                db.CloseConnection(ref sqlConn);

                foreach (DataRow row in dt.Rows)
                {
                    string NomorDraft = util.GetStringValue(row, "NomorDraft");
                    db.OpenConnection(ref sqlConn);
                    db.cmd.CommandText = "usp_GeneratePembiayaan";
                    db.cmd.CommandType = CommandType.StoredProcedure;
                    db.cmd.Parameters.Clear();
                    db.AddInParameter(db.cmd, "NomorDraft", NomorDraft);
                    reader = db.cmd.ExecuteReader();
                    dtDetail.Load(reader);
                    db.CloseDataReader(reader);
                    db.CloseConnection(ref sqlConn);

                    string Officer = "";
                    string txt_file_name = "PEMBIAYAAN_" + DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("HHmmss") + "_" + util.GetStringValue(row, "TotalRow") + ".txt"; //kurang total row
                    string file_output_url = Path.Combine(Environment.CurrentDirectory, "Output") + "\\TXT\\Pembiayaan\\" + txt_file_name;

                    #region Write Txt
                    foreach (DataRow r in dtDetail.Rows)
                    {
                        Officer = util.GetStringValue(row, "GeneratedBy");

                        string Txt = util.GetStringValue(r, "Currency");
                        using (StreamWriter writer = new StreamWriter(file_output_url, true))
                        {
                            Txt += "|"; // Proposal Date
                            Txt += "|" + util.GetStringValue(r, "PlafonRekomendasi");
                            Txt += "|"; //Maximum Total
                            Txt += "|" + util.GetStringValue(r, "RencanaCair");
                            Txt += "|" + util.GetStringValue(r, "ExpiryDate");
                            Txt += "|"; //Review Frekuency
                            Txt += "|"; //Liab Group
                            Txt += "|"; //Notes
                            Txt += "|"; //Cust Name
                            Txt += "|" + util.GetStringValue(r, "Currency");
                            Txt += "|" + util.GetStringValue(r, "PlafonRekomendasi");
                            Txt += "|" + util.GetStringValue(r, "RencanaCair");
                            Txt += "|" + util.GetStringValue(r, "PlafonRekomendasi");
                            Txt += "|"; //EXPIRY.DATE
                            Txt += "|"; //T.GROUP.ID
                            Txt += "|"; //REVIEW.FREQUENCY
                            Txt += "|"; //NOTES
                            Txt += "|"; //ORIENTATION
                            Txt += "|"; //PRODUCT.CHAR
                            Txt += "|"; //CLASS.OF.CREDIT
                            Txt += "|"; //PROJECT.LOCATE
                            Txt += "|"; // TYPE.OF.USE
                            Txt += "|"; //ECONOMIC.SECTOR
                            Txt += "|"; //LOANS.CHARC
                            Txt += "|"; //LBUS.TYPE.USE
                            Txt += "|"; //NEW.EXTEND
                            Txt += "|0"; //Plafon BFR
                            Txt += "|"; //PLAOB.TYPE
                            Txt += "|"; //PLAOB.DESC
                            Txt += "|"; //IA.LOAN.CHARC
                            Txt += "|"; //PK.NUMBER
                            Txt += "|"; //FIRST.PK.DATE
                            Txt += "|"; //LAST.PK.NUMBER
                            Txt += "|"; //LAST.PK.DATE
                            Txt += "|"; //BMPK.DIF.VALUE
                            Txt += "|"; //BMPK.DIF.PRCTG
                            Txt += "|"; //BMPK.NOTE
                            Txt += "|"; //CATEGORY.BR
                            Txt += "|"; //SHORT.DESC
                            Txt += "|"; //DESCRIPTION
                            Txt += "|" + util.GetStringValue(r, "CIF");
                            Txt += "|" + util.GetStringValue(r, "Currency");
                            Txt += "|"; // CUST.LIMIT
                            Txt += "|"; //SUPPLIER.ID
                            Txt += "|"; //SUPPLIER.NAME
                            Txt += "|"; //SUPPLIER.ACCT
                            Txt += "|" + util.GetStringValue(r, "PlafonRekomendasi");
                            Txt += "|" + util.GetStringValue(r, "DownPayment");
                            Txt += "|" + util.GetStringValue(r, "AssetQty");
                            Txt += "|"; //HPP ASSET
                            Txt += "|" + util.GetStringValue(r, "CIF");
                            Txt += "|" + util.GetStringValue(r, "Currency");
                            Txt += "|" + util.GetStringValue(r, "ProdType");
                            Txt += "|"; //IAR.REF
                            Txt += "|"; //AMOUNT
                            Txt += "|"; //LIMIT.REFERENCE
                            Txt += "|" + util.GetStringValue(r, "TenorRekomendasi");
                            Txt += "|" + util.GetStringValue(r, "SchdType");
                            Txt += "|"; // FILE.NAME
                            Txt += "|"; //CUST.ACCT
                            Txt += "|"; //PRIN.LIQ.ACCT
                            Txt += "|"; //INT.LIQ.ACCT
                            Txt += "|" + util.GetStringValue(r, "WakalahFlag");
                            Txt += "|" + util.GetStringValue(r, "SingleMulti");
                            Txt += "|"; // TIER.PERIOD
                            Txt += "|"; // TIER.RATE
                            Txt += "|" + util.GetStringValue(r, "MarginRekomendasi");
                            Txt += "|"; // GRC.DURATION
                            Txt += "|"; // MRG.AMT
                            Txt += "|"; // CHARGE.CODE
                            Txt += "|"; // CHRG.AMT
                            Txt += "|"; // TOT.CHRG.AMT
                            Txt += "|"; // CHRG.LIQ.ACCT
                            Txt += "|" + util.GetStringValue(r, "AGNFlag");

                            Txt += "|"; //Coll Code
                            Txt += "|"; //Percent Alloc
                            Txt += "|" + util.GetStringValue(r, "StatusPembiaya");
                            Txt += "|" + util.GetStringValue(r, "ClassOfCredit");
                            Txt += "|" + util.GetStringValue(r, "PortfolioCateg");

                            Txt += "|"; //CONDITION
                            Txt += "|"; //CONDITION.DATE
                            Txt += "|"; //FREQ.INS
                            Txt += "|"; //PLAOB.DESC
                            Txt += "|"; //DEBTOR.PROB

                            Txt += "|" + util.GetStringValue(r, "PendapatanBersihUsaha");
                            Txt += "|"; //STAGNANT.DATE
                            Txt += "|"; //STAGNANT.REASON
                            Txt += "|" + util.GetStringValue(r, "TypeOfUse");
                            Txt += "|" + util.GetStringValue(r, "LoansChar");
                            Txt += "|" + util.GetStringValue(r, "RencanaCair");
                            Txt += "|" + util.GetStringValue(r, "JatuhTempo");
                            Txt += "|" + util.GetStringValue(r, "NomorAkad");
                            Txt += "|" + util.GetStringValue(r, "NomorAkad");
                            Txt += "|" + util.GetStringValue(r, "Dati2");
                            Txt += "|";
                            Txt += "|";
                            Txt += "|" + util.GetStringValue(r, "InstDate");

                            writer.WriteLine(Txt);

                        }
                    }
                    #endregion

                    new SharePointHelper().UploadFileToDocLib(file_output_url, "Pembiayaan");
                    EmailTxt eml = new EmailTxt();
                    eml.Email = util.GetStringValue(row, "GeneratedBy");
                    eml.file_attachment = file_output_url;
                    eml.NomorDraft = NomorDraft;
                    list.Add(eml);
                }
                return list;
            }
            catch (Exception ex)
            {
                db.CloseConnection(ref sqlConn);
                throw ex;
            }
        }

        public List<EmailTxt> GenerateTxtCIF()
        {
            DataTable dt = new DataTable();
            DataTable dtDetail = new DataTable();
            List<EmailTxt> list = new List<EmailTxt>();
            try
            {
                db.OpenConnection(ref sqlConn);
                db.cmd.CommandText = "usp_List_UnGenerate_CIF_Pembiayaan_PerAgent";
                db.cmd.CommandType = CommandType.StoredProcedure;
                db.cmd.Parameters.Clear();
                db.AddInParameter(db.cmd, "GenerateType", "CIF");
                reader = db.cmd.ExecuteReader();
                dt.Load(reader);
                db.CloseDataReader(reader);
                db.CloseConnection(ref sqlConn);

                foreach (DataRow row in dt.Rows)
                {
                    string NomorDraft = util.GetStringValue(row, "NomorDraft");
                    string txt_file_name = "CIF_" + DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("HH:mm:ss") + "_" + ".txt"; //kurang total row

                    string file_output_url = Path.Combine(Environment.CurrentDirectory, "Output") + "\\TXT\\CIF\\" + txt_file_name;

                    db.OpenConnection(ref sqlConn);
                    db.cmd.CommandText = "usp_GenerateCIF";
                    db.cmd.CommandType = CommandType.StoredProcedure;
                    db.AddInParameter(db.cmd, "NomorDraft", NomorDraft);

                    reader = db.cmd.ExecuteReader();

                    dtDetail.Load(reader);

                    db.CloseDataReader(reader);
                    db.CloseConnection(ref sqlConn);

                    #region Write Txt

                    foreach (DataRow rowDetail in dtDetail.Rows)
                    {

                        using (StreamWriter writer = new StreamWriter(file_output_url, true))
                        {
                            string Txt = "|" + util.GetStringValue(rowDetail, "CIF");
                            Txt += "|" + util.GetStringValue(rowDetail, "MNEMONIC");
                            Txt += "|" + util.GetStringValue(rowDetail, "CUSTTITLE1");
                            Txt += "|" + util.GetStringValue(rowDetail, "NamaSesuaiKTP");
                            Txt += "|" + util.GetStringValue(rowDetail, "CUSTTITLE2");
                            Txt += "|" + util.GetStringValue(rowDetail, "NamaSesuaiKTP");
                            Txt += "|" + util.GetStringValue(rowDetail, "Name2");
                            Txt += "|" + util.GetStringValue(rowDetail, "Alias") + "|" + util.GetStringValue(rowDetail, "JenisKelamin");
                            Txt += "|" + util.GetStringValue(rowDetail, "TempatLahir") + "|" + util.GetStringValue(rowDetail, "TanggalLahir");
                            Txt += "|" + util.GetStringValue(rowDetail, "NamaIbuGadisKandung") + "|" + util.GetStringValue(rowDetail, "LegalType");
                            Txt += "|" + util.GetStringValue(rowDetail, "NoKTP") + "|" + util.GetStringValue(rowDetail, "MasaBerlaku");
                            Txt += "|" + util.GetStringValue(rowDetail, "Reside") + "|" + util.GetStringValue(rowDetail, "Nationality");
                            Txt += "|" + util.GetStringValue(rowDetail, "Taxable") + "|" + util.GetStringValue(rowDetail, "NPWP");
                            Txt += "|" + util.GetStringValue(rowDetail, "Agama") + "|" + util.GetStringValue(rowDetail, "StatusPerkawinan");
                            Txt += "|" + util.GetStringValue(rowDetail, "PendidikanTerakhir");
                            Txt += "|" + util.GetStringValue(rowDetail, "EDUCATIONOTHER");
                            Txt += "|" + util.GetStringValue(rowDetail, "Sector");
                            Txt += "|" + util.GetStringValue(rowDetail, "Industry");
                            Txt += "|" + util.GetStringValue(rowDetail, "Target");
                            Txt += "|" + util.GetStringValue(rowDetail, "NamaPetugas") + "|" + util.GetStringValue(rowDetail, "CustType");
                            Txt += "|" + util.GetStringValue(rowDetail, "Language") + "|" + util.GetStringValue(rowDetail, "Alamat_KTP");
                            Txt += "|" + util.GetStringValue(rowDetail, "Address");
                            Txt += "|" + util.GetStringValue(rowDetail, "RT_RW") + "|" + util.GetStringValue(rowDetail, "Provinsi_KTP");
                            Txt += "|" + util.GetStringValue(rowDetail, "Kecamatan_KTP") + "|" + util.GetStringValue(rowDetail, "Kelurahan_KTP");
                            Txt += "|" + util.GetStringValue(rowDetail, "Residence") + "|" + util.GetStringValue(rowDetail, "Kabupaten_KTP");
                            Txt += "|" + util.GetStringValue(rowDetail, "KodePos_KTP") + "|" + util.GetStringValue(rowDetail, "StatusMilikTempatTinggal");
                            Txt += "|" + util.GetStringValue(rowDetail, "OTHResStatus") + "|" + util.GetStringValue(rowDetail, "NoTelp");
                            Txt += "|" + util.GetStringValue(rowDetail, "OFFPhone");
                            Txt += "|" + util.GetStringValue(rowDetail, "Fax1");
                            Txt += "|" + util.GetStringValue(rowDetail, "NoHP");
                            Txt += "|" + util.GetStringValue(rowDetail, "Email1");
                            Txt += "|" + util.GetStringValue(rowDetail, "ADDRType");
                            Txt += "|" + util.GetStringValue(rowDetail, "ADDRStreet");
                            Txt += "|" + util.GetStringValue(rowDetail, "ADDRRTRW");
                            Txt += "|" + util.GetStringValue(rowDetail, "ADDRProvince");
                            Txt += "|" + util.GetStringValue(rowDetail, "ADDRSUBBRTWN");
                            Txt += "|" + util.GetStringValue(rowDetail, "MUNICIPAL");
                            Txt += "|" + util.GetStringValue(rowDetail, "COUNTRY");
                            Txt += "|" + util.GetStringValue(rowDetail, "DISTRICT");
                            Txt += "|" + util.GetStringValue(rowDetail, "POSTCODE");
                            Txt += "|" + util.GetStringValue(rowDetail, "Pekerjaan");
                            Txt += "|" + util.GetStringValue(rowDetail, "EmploymentStatus");
                            Txt += "|" + util.GetStringValue(rowDetail, "OCCUPATION");
                            Txt += "|" + util.GetStringValue(rowDetail, "KodeSektorEkonomi");
                            Txt += "|" + util.GetStringValue(rowDetail, "EMPLOYERSName");
                            Txt += "|" + util.GetStringValue(rowDetail, "EmployersAdd");
                            Txt += "|" + util.GetStringValue(rowDetail, "EmploymentStart");
                            Txt += "|" + util.GetStringValue(rowDetail, "FundProvName");
                            Txt += "|" + util.GetStringValue(rowDetail, "FundProvJob");
                            Txt += "|" + util.GetStringValue(rowDetail, "FundProvAddr");
                            Txt += "|" + util.GetStringValue(rowDetail, "FundProvPhone");

                            Txt += "|" + util.GetStringValue(rowDetail, "FundSource");
                            Txt += "|" + util.GetStringValue(rowDetail, "OTHFundSource");
                            Txt += "|" + util.GetStringValue(rowDetail, "FundSourceAMT");
                            Txt += "|" + util.GetStringValue(rowDetail, "OTHAcctType");
                            Txt += "|" + util.GetStringValue(rowDetail, "OTHAcctNo");
                            Txt += "|" + util.GetStringValue(rowDetail, "OTHACBranch");
                            Txt += "|" + util.GetStringValue(rowDetail, "OTHACBNKName");
                            Txt += "|" + util.GetStringValue(rowDetail, "OTHACOpened");
                            Txt += "|" + util.GetStringValue(rowDetail, "OTHRemarks");
                            Txt += "|" + util.GetStringValue(rowDetail, "ContactName");
                            Txt += "|" + util.GetStringValue(rowDetail, "ContactStreet");
                            Txt += "|" + util.GetStringValue(rowDetail, "ContactHomtel");
                            Txt += "|" + util.GetStringValue(rowDetail, "ContactRelCus");
                            Txt += "|" + util.GetStringValue(rowDetail, "NoDebitTrans");
                            Txt += "|" + util.GetStringValue(rowDetail, "ValueDRTrans");
                            Txt += "|" + util.GetStringValue(rowDetail, "NoCreditTrans");
                            Txt += "|" + util.GetStringValue(rowDetail, "ValueCRTrans");
                            Txt += "|" + util.GetStringValue(rowDetail, "HighRisk");
                            Txt += "|" + util.GetStringValue(rowDetail, "GuarantorCode");
                            Txt += "|" + util.GetStringValue(rowDetail, "SidRelatiBank");
                            Txt += "|" + util.GetStringValue(rowDetail, "DINNumber");
                            Txt += "|" + util.GetStringValue(rowDetail, "BMPKViolation");
                            Txt += "|" + util.GetStringValue(rowDetail, "BMPKExceeding");
                            Txt += "|" + util.GetStringValue(rowDetail, "LBU_Cust_Type");
                            Txt += "|" + util.GetStringValue(rowDetail, "CustomerRating");
                            Txt += "|" + util.GetStringValue(rowDetail, "CustomerSince");
                            Txt += "|" + util.GetStringValue(rowDetail, "UploadCompany");
                            Txt += "|" + util.GetStringValue(rowDetail, "RESStatus");

                            Txt += "|" + util.GetStringValue(rowDetail, "RESYEAR") + "|" + util.GetStringValue(rowDetail, "RESMONTH");
                            Txt += "|" + util.GetStringValue(rowDetail, "TotalEmployee") + "|" + util.GetStringValue(rowDetail, "RelatiBank");
                            Txt += "|" + util.GetStringValue(rowDetail, "TotalLiability") + "|" + util.GetStringValue(rowDetail, "AttStatus");
                            Txt += "|" + util.GetStringValue(rowDetail, "RelationCode");
                            Txt += "|" + util.GetStringValue(rowDetail, "RelCustomer");
                            Txt += "|" + util.GetStringValue(rowDetail, "PortfolioCateg");
                            Txt += "|" + util.GetStringValue(rowDetail, "AddrPhoneArea");
                            Txt += "|" + util.GetStringValue(rowDetail, "NoHpPetugas");
                            Txt += "|" + util.GetStringValue(rowDetail, "LLC");
                            Txt += "|" + util.GetStringValue(rowDetail, "NamaBank");
                            Txt += "|" + util.GetStringValue(rowDetail, "NomorRekening");
                            Txt += "|" + util.GetStringValue(rowDetail, "NamaPemilikRekening");
                            Txt += "|" + util.GetStringValue(rowDetail, "Fatca");
                            Txt += "|" + util.GetStringValue(rowDetail, "TempatLahir");
                            Txt += "|" + util.GetStringValue(rowDetail, "AgentCode");

                            writer.WriteLine(Txt);

                        }
                    }
                    #endregion

                    new SharePointHelper().UploadFileToDocLib(file_output_url, "CIF");
                    EmailTxt eml = new EmailTxt();
                    eml.Email = util.GetStringValue(row, "GeneratedBy");
                    eml.file_attachment = file_output_url;
                    eml.NomorDraft = NomorDraft;
                    list.Add(eml);

                }
                return list;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
