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
    public partial class UploadRHList : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            string FilePath = ConfigurationManager.AppSettings["FilePath"].ToString();
            string filename = string.Empty;

            if(FileUploadToServer.HasFile)
            {
                try
                {
                    string[] allowedFile = { ".xls", ".xlsx" };
                    string FileExt = System.IO.Path.GetExtension(FileUploadToServer.PostedFile.FileName);

                    bool isValidFile = allowedFile.Contains(FileExt);
                    if(!isValidFile)
                    {
                        lblMsg.ForeColor = System.Drawing.Color.Red;
                        lblMsg.Text = "Please upload only Excel format file.";
                    }
                    else
                    {
                        int FileSize = FileUploadToServer.PostedFile.ContentLength;
                        if(FileSize <= 5242880) //5242880 bytes = 5MB
                        {
                            filename = Path.GetFileName(Server.MapPath(FileUploadToServer.FileName));
                            FileUploadToServer.SaveAs(Server.MapPath(FilePath) + filename);

                            string filePath = Server.MapPath(FilePath) + filename;

                            OleDbConnection con = null;
                            if(FileExt == ".xls")
                            {
                                con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=Excel 8.0;");
                            }
                            else if(FileExt == ".xlsx")
                            {
                                con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=Excel 12.0;");
                            }
                            con.Open();

                            DataTable dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            string getExcelSheetName = dt.Rows[0]["Table_Name"].ToString();

                            OleDbCommand ExcelCommand = new OleDbCommand(@"SELECT F1,F2,F3 FROM [" + getExcelSheetName + @"]", con);
                            OleDbDataAdapter ExcelAdapter = new OleDbDataAdapter(ExcelCommand);
                            DataSet ExcelDataSet = new DataSet();
                            ExcelAdapter.Fill(ExcelDataSet);
                            con.Close();

                            GridView1.DataSource = ExcelDataSet;
                            GridView1.DataBind();
                        }
                        else
                        {
                            lblMsg.Text = "Attached file size should not be greater than 5 MB!";
                        }
                    }
                }
                catch(Exception ex)
                {
                    lblMsg.Text = "Error occured while uploading: " + ex.Message;
                }
            }
            else
            {
                lblMsg.Text = "Please select a file to upload.";
            }
        }
    }
}