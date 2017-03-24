using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.OleDb;
using System.Configuration;
using System.IO;
using System.Data;

namespace UploadHoliday
{
    public partial class UploadRH : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            try
            {
                if(UploadRHFile.UploadedFiles.Count > 0)
                {
                    string filename = "~/File/" + UploadRHFile.UploadedFiles[0].GetName();
                    if(UploadRHFile.UploadedFiles[0].GetExtension()==".xls"||UploadRHFile.UploadedFiles[0].GetExtension()==".xlsx")
                    {
                        if(UploadRHFile.UploadedFiles[0].GetName()!="")
                        {
                                string filePath = Server.MapPath(filename);
                                string ConnectionString = "";
                                ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                                   "Data Source=" + filePath + ";" +
                                                   "Extended Properties='Excel 8.0;HDR=NO;IMEX=1'";
                                using (OleDbConnection cn = new OleDbConnection(ConnectionString))
                                {
                                    cn.Close();
                                    cn.Open();
                                    DataTable dbSchema = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                                    if (dbSchema == null || dbSchema.Rows.Count < 1)
                                    {
                                        throw new Exception("Error: Could not determine the name of the first worksheet.");
                                    }
                                    string WorkSheetName = dbSchema.Rows[0]["TABLE_NAME"].ToString();
                                    OleDbDataAdapter da = new OleDbDataAdapter("SELECT F1,F2,F3 FROM [" + WorkSheetName + "]", cn);
                                    DataTable dtFromExcel = new DataTable(WorkSheetName);
                                    da.Fill(dtFromExcel);
                                    cn.Close();

                                GridViewData.DataSource = dtFromExcel;
                                GridViewData.DataBind();

                                }
                        }
                        else
                        {
                            lblMessage.Text = "Please select a file to upload.";
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                lblMessage.Text = "Error occured while uploading: " + ex.Message;
            }
        }
    }
}