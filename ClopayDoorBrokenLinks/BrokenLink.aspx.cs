using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Net;
using NPOI;
using NPOI.HSSF.UserModel;
using System.Data;
using System.IO;
using System.Data.SqlClient;

namespace ClopayDoorBrokenLinks
{
    public partial class BrokenLink : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnResidential_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dtLinksData = GetData("RESIDENTIAL");

                DataTable dtBrokenLinksData = new DataTable();
                dtBrokenLinksData.Columns.Add("FLApproval");
                dtBrokenLinksData.Columns.Add("FLDrawing");
                dtBrokenLinksData.Columns.Add("TxApproval");
                dtBrokenLinksData.Columns.Add("TDIDrawing");
                dtBrokenLinksData.Columns.Add("IBCDrawing");
                dtBrokenLinksData.Columns.Add("DadeApproval");

                dtBrokenLinksData.Columns.Add("Title");
                dtBrokenLinksData.Columns.Add("IdealDealerModelNumbers");
                dtBrokenLinksData.Columns.Add("HolmesModelNumbers");
                dtBrokenLinksData.Columns.Add("ClopayRetailModelNumbers");
                dtBrokenLinksData.Columns.Add("ClopayDealerModelNumbers");

                DataRow dr = null;

                for (int i = 0; i < dtLinksData.Rows.Count; i++)
                {
                    var FLApproval = IsUrlAvailable(dtLinksData.Rows[i]["FLApproval"].ToString());
                    var FLDrawing = IsUrlAvailable(dtLinksData.Rows[i]["FLDrawing"].ToString());

                    var TxApproval = IsUrlAvailable(dtLinksData.Rows[i]["TxApproval"].ToString());
                    var TDIDrawing = IsUrlAvailable(dtLinksData.Rows[i]["TDIDrawing"].ToString());
                    var IBCDrawing = IsUrlAvailable(dtLinksData.Rows[i]["IBCDrawing"].ToString());
                    var DadeApproval = IsUrlAvailable(dtLinksData.Rows[i]["DadeApproval"].ToString());  

                    dr = dtBrokenLinksData.NewRow(); // have new row on each iteration
                    if (!FLApproval)
                        dr["FLApproval"] = dtLinksData.Rows[i]["FLApproval"].ToString();
                    if (!FLDrawing)
                        dr["FLDrawing"] = dtLinksData.Rows[i]["FLDrawing"].ToString();

                    if (!TxApproval)
                        dr["TxApproval"] = dtLinksData.Rows[i]["TxApproval"].ToString();

                    if (!TDIDrawing)
                        dr["TDIDrawing"] = dtLinksData.Rows[i]["TDIDrawing"].ToString();

                    if (!IBCDrawing)
                        dr["IBCDrawing"] = dtLinksData.Rows[i]["IBCDrawing"].ToString();

                    if (!DadeApproval)
                        dr["DadeApproval"] = dtLinksData.Rows[i]["DadeApproval"].ToString();

                    //Models
                    dr["Title"] = dtLinksData.Rows[i]["Title"].ToString();
                    dr["IdealDealerModelNumbers"] = dtLinksData.Rows[i]["IdealDealerModelNumbers"].ToString();
                    dr["HolmesModelNumbers"] = dtLinksData.Rows[i]["HolmesModelNumbers"].ToString();
                    dr["ClopayRetailModelNumbers"] = dtLinksData.Rows[i]["ClopayRetailModelNumbers"].ToString();
                    dr["ClopayDealerModelNumbers"] = dtLinksData.Rows[i]["ClopayDealerModelNumbers"].ToString();

                    if (dr["FLApproval"].ToString() != "" || dr["FLDrawing"].ToString() != "" || dr["TxApproval"].ToString() != ""
                        || dr["TxApproval"].ToString() != "" || dr["TDIDrawing"].ToString() != "" || dr["IBCDrawing"].ToString() != ""
                        || dr["DadeApproval"].ToString() != "")
                    {
                        dtBrokenLinksData.Rows.Add(dr);
                    }
                }

                if (dtBrokenLinksData.Rows.Count > 0)
                {
                    DataTableToExcel(dtBrokenLinksData, "BrokenLinks_Residential_Data_" + DateTime.Now.ToString());
                }
                else
                {
                    lblmsg.Text = "No Data to download";
                }
            }
            catch (Exception)
            {

                throw;
            }

        }

        protected void btnModel_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dtLinksData = GetData("MODEL");

                DataTable dtBrokenLinksData = new DataTable();
                dtBrokenLinksData.Columns.Add("Brochure");
                dtBrokenLinksData.Columns.Add("InstallationManual");
                dtBrokenLinksData.Columns.Add("Specs");
                dtBrokenLinksData.Columns.Add("Warranty");
                DataRow dr = null;

                for (int i = 0; i < dtLinksData.Rows.Count; i++)
                {
                    var Brochure = IsUrlAvailable(dtLinksData.Rows[i]["Brochure"].ToString());
                    var InstallationManual = IsUrlAvailable(dtLinksData.Rows[i]["InstallationManual"].ToString());
                    var Specs = IsUrlAvailable(dtLinksData.Rows[i]["Specs"].ToString());
                    var Warranty = IsUrlAvailable(dtLinksData.Rows[i]["Warranty"].ToString());

                    dr = dtBrokenLinksData.NewRow(); // have new row on each iteration
                    if (!Brochure)
                        dr["Brochure"] = dtLinksData.Rows[i]["Brochure"].ToString();
                    if (!InstallationManual)
                        dr["InstallationManual"] = dtLinksData.Rows[i]["InstallationManual"].ToString();
                    if (!Specs)
                        dr["Specs"] = dtLinksData.Rows[i]["Specs"].ToString();
                    if (!Warranty)
                        dr["Warranty"] = dtLinksData.Rows[i]["Warranty"].ToString();

                    if (dr["Brochure"].ToString() != "" || dr["InstallationManual"].ToString() != "" || dr["Specs"].ToString() != "" || dr["Warranty"].ToString() != "")
                        dtBrokenLinksData.Rows.Add(dr);
                }

                if (dtBrokenLinksData.Rows.Count > 0)
                {
                    DataTableToExcel(dtBrokenLinksData, "BrokenLinks_Model_Data_" + DateTime.Now.ToString());
                }
                else
                {
                    lblmsg.Text = "No Data to download";
                }
            }
            catch (Exception)
            {

                throw;
            }

        }

        protected void btnCommercial_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dtLinksData = GetData("COMMERCIAL");

                DataTable dtBrokenLinksData = new DataTable();
                dtBrokenLinksData.Columns.Add("TxDrawing");
                dtBrokenLinksData.Columns.Add("TxDrawing2");
                dtBrokenLinksData.Columns.Add("TxApproval");
                dtBrokenLinksData.Columns.Add("TDIListedStandard");
                dtBrokenLinksData.Columns.Add("TDIListedImpactResistant");
                dtBrokenLinksData.Columns.Add("IBCDrawing");
                dtBrokenLinksData.Columns.Add("FLApproval");
                dtBrokenLinksData.Columns.Add("Drawing");
                dtBrokenLinksData.Columns.Add("DadeListed");
                dtBrokenLinksData.Columns.Add("ApprovalDrawing");

                dtBrokenLinksData.Columns.Add("Title");
                dtBrokenLinksData.Columns.Add("IdealDealerModelNumbers");
                dtBrokenLinksData.Columns.Add("HolmesModelNumbers");
                dtBrokenLinksData.Columns.Add("ClopayDealerModelNumbers");

                DataRow dr = null;

                for (int i = 0; i < dtLinksData.Rows.Count; i++)
                {
                    var TxDrawing = IsUrlAvailable(dtLinksData.Rows[i]["TxDrawing"].ToString());
                    var TxDrawing2 = IsUrlAvailable(dtLinksData.Rows[i]["TxDrawing2"].ToString());
                    var TxApproval = IsUrlAvailable(dtLinksData.Rows[i]["TxApproval"].ToString());
                    var TDIListedStandard = IsUrlAvailable(dtLinksData.Rows[i]["TDIListedStandard"].ToString());
                    var TDIListedImpactResistant = IsUrlAvailable(dtLinksData.Rows[i]["TDIListedImpactResistant"].ToString());
                    var IBCDrawing = IsUrlAvailable(dtLinksData.Rows[i]["IBCDrawing"].ToString());
                    var FLApproval = IsUrlAvailable(dtLinksData.Rows[i]["FLApproval"].ToString());
                    var Drawing = IsUrlAvailable(dtLinksData.Rows[i]["Drawing"].ToString());
                    var DadeListed = IsUrlAvailable(dtLinksData.Rows[i]["DadeListed"].ToString());
                    var ApprovalDrawing = IsUrlAvailable(dtLinksData.Rows[i]["ApprovalDrawing"].ToString());

                    dr = dtBrokenLinksData.NewRow(); // have new row on each iteration
                    if (!TxDrawing)
                        dr["TxDrawing"] = dtLinksData.Rows[i]["TxDrawing"].ToString();
                    if (!TxDrawing2)
                        dr["TxDrawing2"] = dtLinksData.Rows[i]["TxDrawing2"].ToString();
                    if (!TxApproval)
                        dr["TxApproval"] = dtLinksData.Rows[i]["TxApproval"].ToString();
                    if (!TDIListedStandard)
                        dr["TDIListedStandard"] = dtLinksData.Rows[i]["TDIListedStandard"].ToString();
                    if (!TDIListedImpactResistant)
                        dr["TDIListedImpactResistant"] = dtLinksData.Rows[i]["TDIListedImpactResistant"].ToString();
                    if (!IBCDrawing)
                        dr["IBCDrawing"] = dtLinksData.Rows[i]["IBCDrawing"].ToString();
                    if (!FLApproval)
                        dr["FLApproval"] = dtLinksData.Rows[i]["FLApproval"].ToString();
                    if (!Drawing)
                        dr["Drawing"] = dtLinksData.Rows[i]["Drawing"].ToString();
                    if (!DadeListed)
                        dr["DadeListed"] = dtLinksData.Rows[i]["DadeListed"].ToString();
                    if (!ApprovalDrawing)
                        dr["ApprovalDrawing"] = dtLinksData.Rows[i]["ApprovalDrawing"].ToString();

                    //Models
                    dr["Title"] = dtLinksData.Rows[i]["Title"].ToString();
                    dr["HolmesModelNumbers"] = dtLinksData.Rows[i]["HolmesModelNumbers"].ToString();
                    dr["IdealDealerModelNumbers"] = dtLinksData.Rows[i]["IdealDealerModelNumbers"].ToString();                                        
                    dr["ClopayDealerModelNumbers"] = dtLinksData.Rows[i]["ClopayDealerModelNumbers"].ToString();


                    if (dr["TxDrawing"].ToString() != "" || dr["TxDrawing2"].ToString() != "" || dr["TxApproval"].ToString() != ""
                        || dr["TDIListedStandard"].ToString() != "" || dr["TDIListedImpactResistant"].ToString() != "" || dr["IBCDrawing"].ToString() != ""
                        || dr["FLApproval"].ToString() != "" || dr["Drawing"].ToString() != "" || dr["DadeListed"].ToString() != "" || dr["ApprovalDrawing"].ToString() != "")
                    {
                        dtBrokenLinksData.Rows.Add(dr);
                    }
                }

                if (dtBrokenLinksData.Rows.Count > 0)
                {
                    DataTableToExcel(dtBrokenLinksData, "BrokenLinks_Commercial_Data_" + DateTime.Now.ToString());
                }
                else
                {
                    lblmsg.Text = "No Data to download";
                }
            }
            catch (Exception)
            {

                throw;
            }


        }

        public static DataTable GetData(string type)
        {
            string dbConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["ClopayDoorPortal"].ConnectionString;
            using (SqlConnection cnClopayDB = new SqlConnection(dbConnectionString))
            {

                using (SqlCommand cmdBrokenLinksData = new SqlCommand())
                {
                    DataTable dt = new DataTable();
                    cmdBrokenLinksData.CommandText = "dbo.GetBrokenLinksData";
                    cmdBrokenLinksData.CommandType = CommandType.StoredProcedure;
                    cmdBrokenLinksData.Connection = cnClopayDB;
                    cmdBrokenLinksData.Parameters.AddWithValue("@type", type);
                    SqlDataAdapter sda = new SqlDataAdapter();
                    try
                    {
                        cnClopayDB.Open();
                        sda.SelectCommand = cmdBrokenLinksData;
                        sda.Fill(dt);
                    }
                    catch (IOException ex)
                    {

                        return null;
                    }
                    finally
                    {
                        cnClopayDB.Close();
                        sda.Dispose();
                        cnClopayDB.Dispose();
                    }

                    return dt;
                }
            }
        }

        public bool IsUrlAvailable(string url)
        {
            try
            {
                if (url.Contains("http") || url.Contains("https"))
                {
                    HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);

                    using (HttpWebResponse rsp = (HttpWebResponse)req.GetResponse())
                    {
                        if (rsp.StatusCode == HttpStatusCode.OK)
                        {
                            return true;
                        }
                    }
                }
            }
            catch (WebException)
            {

            }

            return false;
        }

        private void DataTableToExcel(DataTable dt, String fileName)
        {

            //Make a new npoi workbook
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();
            //Here I am making sure that I am giving the file name the right extension:
            string filename = "";
            if (fileName.EndsWith(".xls"))
            {
                filename = fileName;
            }
            else
            {
                filename = fileName + ".xls";
            }

            //This starts the dialogue box that allows the user to download the file
            System.Web.HttpResponse Response = System.Web.HttpContext.Current.Response;
            Response.ContentType = "application/vnd.ms-excel";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();

            //make a new sheet – name it any excel-compliant string you want
            HSSFSheet sheet1 = (HSSFSheet)hssfworkbook.CreateSheet("Sheet1");
            //make a header row
            var row1 = sheet1.CreateRow(0);
            //Puts in headers (these are table row headers, omit if you
            //just need a straight data dump
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                var cell = row1.CreateCell(j);
                String columnName = dt.Columns[j].ToString();
                cell.SetCellValue(columnName);
            }

            //loops through data
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                var row = sheet1.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    var cell = row.CreateCell(j);
                    String columnName = dt.Columns[j].ToString();
                    cell.SetCellValue(dt.Rows[i][columnName].ToString());
                }
            }
            //writing the data to binary from memory
            Response.BinaryWrite(WriteToStream(hssfworkbook).GetBuffer());
            // Response.End(false);
            HttpContext.Current.Response.Flush(); // Sends all currently buffered output to the client.
            HttpContext.Current.Response.SuppressContent = true;  // Gets or sets a value indicating whether to send HTTP content to the client.
            HttpContext.Current.ApplicationInstance.CompleteRequest();

        }

        static MemoryStream WriteToStream(HSSFWorkbook hssfworkbook)
        {
            //Write the stream data of workbook to the root directory
            MemoryStream file = new MemoryStream();
            hssfworkbook.Write(file);
            return file;
        }
    }
}