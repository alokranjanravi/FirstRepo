using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.UI.WebControls;
using System.IO.Compression;
using System.Web.UI.HtmlControls;
using System.Data;
using System.Data.OleDb;

using System.Xml;
using Org.BouncyCastle.Asn1.Cmp;
using System.Data.Common;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Reflection;
using static System.Net.WebRequestMethods;
using File = System.IO.File;
using Path = System.IO.Path;
using System.Web.UI;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.Asn1;

namespace MsgE
{
    public partial class WebForm1 : System.Web.UI.Page
    {

        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void btnExtract_Click(object sender, EventArgs e)
        {

            //Section to create Directory & Sub Directory
            lblMessage.Text = string.Empty;
            lblExtract.Text = string.Empty;
            lblFailed.Text = string.Empty;
            lblTotal.Text = string.Empty;
            string root = @"C:\Unzip";
            string subdir = @"C:\Unzip\01-HTML_Files";
            string subdir_2 = @"C:\Unzip\04-Zip_File";
            string subdir_3 = @"C:\Unzip\05-Email_Files";
            string subdir_4 = @"C:\Unzip\EmailLog";
            string msgPath = @"C:\Unzip\05-Email_Files\Done\";
            string zippedPath = @"C:\\Unzip\\04-Zip_File\Done\";
           
            // If directory does not exist, create it. 
            if (!Directory.Exists(root))
            {
                Directory.CreateDirectory(root);
            }
            // Create a sub directory
            if (!Directory.Exists(subdir))
            {
                Directory.CreateDirectory(subdir);
            }
            if (!Directory.Exists(subdir_2))
            {
                Directory.CreateDirectory(subdir_2);
            }
            if (!Directory.Exists(subdir_3))
            {
                Directory.CreateDirectory(subdir_3);
            }
            if (!Directory.Exists(subdir_4))
            {
                Directory.CreateDirectory(subdir_4);
            }
            if (!Directory.Exists(msgPath))
            {
                Directory.CreateDirectory(msgPath);
            }
            if (!Directory.Exists(zippedPath))
            {
                Directory.CreateDirectory(zippedPath);
            }
          
            //Section to extract .msg files

            _Application outlook;
            outlook = new Microsoft.Office.Interop.Outlook.Application();
            DirectoryInfo DIRINF = new DirectoryInfo("C:\\Unzip\\05-Email_Files");
            List<FileInfo> FINFO = DIRINF.GetFiles("*.msg").ToList();
            List<object> Data = new List<object>();
            int mailcount;
            mailcount = FINFO.Count;
            if (mailcount == 0)
            {
                btnReset.Visible = true;
                lblMail.Visible = true;
                lblMail.Text = "No Emails found !";
            }
            foreach (FileInfo FoundFile in FINFO)
            {
               
                var Name = FoundFile.Name; // Gets the name
                var Path = FoundFile.FullName; // Gets the full path
                var Extension = FoundFile.Extension; // Gets the extension 
                MailItem item = (MailItem)outlook.CreateItemFromTemplate(Path, Type.Missing);
                for (int i = 1; i < item.Attachments.Count + 1; i++)
                {
                    string attpath = @"C:\\Unzip\\04-Zip_File\" + item.Attachments[i].FileName;
                    item.Attachments[i].SaveAsFile(attpath);
                    ExtractZip();
                                
                }
         

            }

        }

        //Section to extract .zip files
       public void ExtractZip()

        {
            string startPath = @"C:\\Unzip\\04-Zip_File\";
            string extractPath = @"C:\\Unzip\\01-HTML_Files\";
            //string zippedPath = @"C:\\Unzip\\04-Zip_File\Done\";
            DirectoryInfo directorySelected = new DirectoryInfo(startPath);
            List<string> fileNames = new List<string>();
            foreach (FileInfo fileInfo in directorySelected.GetFiles("*.zip"))
            {
                fileNames.Add(fileInfo.Name);

            }
            foreach (string name in fileNames)
            {

                string zipfilePath = startPath + "\\" + name;
                using (ZipArchive archive = ZipFile.OpenRead(zipfilePath))
                {
                   foreach (ZipArchiveEntry entry in archive.Entries)
                    {

                        string htmlfile = entry.FullName;
                        string unzipFileName = Path.Combine(extractPath,
                            entry.FullName).Replace("/", "\\");
                        string directoryPath = Path.GetDirectoryName(unzipFileName);
                        if (!Directory.Exists(directoryPath))
                            Directory.CreateDirectory(directoryPath);
                        if (entry.Name == "")
                            continue;
                    }


                }
                //Section to fetch data from Excel       
                using (Ionic.Zip.ZipFile file = new Ionic.Zip.ZipFile(zipfilePath))
                {
                    string ExcelPath = @"C:\Unzip\Mail_pswd.xls";
                    String Con = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelPath + ";Extended Properties=\"Excel 12.0; HDR = Yes; IMEX = 2\"";
                    OleDbConnection obj_Con = new OleDbConnection(Con);
                    obj_Con.Open();
                    string sheet_Name = "Sheet1";
                    string query = String.Format("select * from [{0}$]", sheet_Name);
                    OleDbCommand obj_CmdSelect = new OleDbCommand(query, obj_Con);
                    OleDbDataAdapter obj_Adapter = new OleDbDataAdapter();
                    obj_Adapter.SelectCommand = obj_CmdSelect;
                    DataSet obj_Dataset = new DataSet();
                    obj_Adapter.Fill(obj_Dataset, "Data");
                   //Section to retrieve password wrt filename from Excel file
                    for (int i = 0; i < obj_Dataset.Tables[0].Rows.Count; i++)
                    {
                        
                        DataRow r = obj_Dataset.Tables[0].Rows[i];
                        var filename = obj_Dataset.Tables[0].Rows[i]["FileName"];
                        var filepswd = obj_Dataset.Tables[0].Rows[i]["Password"];
                        try
                        {
                            if (name.Contains((string)filename) == true)
                            {

                                file.Password = filepswd.ToString();
                                file.Encryption = Ionic.Zip.EncryptionAlgorithm.WinZipAes256;
                                file.StatusMessageTextWriter = Console.Out;
                                file.ExtractAll(extractPath);
                            }
                                                    
                        }
                        catch (System.Exception ex)
                        {
                            this.LogError(ex);
                                                 
                        }
                        finally
                        {
                            int mcount = Directory.GetFiles(@"C:\\Unzip\\04-Zip_File\", "*.zip", SearchOption.AllDirectories).Count();
                            lblTotal.Text = "Total Files : " + mcount.ToString(); //Counting total zipped files extracted
                            int cnt = Directory.GetFiles(@"C:\\Unzip\\01-HTML_Files\", "*.html", SearchOption.TopDirectoryOnly).Count();
                            if (cnt > 0)
                            {

                                lblMessage.Text = "File Succesfully Extracted !";

                            }
                            lblExtract.Text = "Success : " + cnt;
                            int xcount = mcount - cnt;
                            lblFailed.Text = "Failed : " + xcount;
                            btnExtract.Enabled = false;
                            btnReset.Visible = true;
                            obj_Con.Close();

                        }

                    }
      

                }
                
               
            }

        }

       private void LogError(System.Exception ex)
        {
            string message = string.Format("Time: {0}", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));
            message += Environment.NewLine;
            message += "-----------------------------------------------------------";
            message += Environment.NewLine;
            message += string.Format("Message: {0}", ex.Message);
            message += Environment.NewLine;
            message += string.Format("StackTrace: {0}", ex.StackTrace);
            message += Environment.NewLine;
            message += string.Format("Source: {0}", ex.Source);
            message += Environment.NewLine;
            message += string.Format("TargetSite: {0}", ex.TargetSite.ToString());
            message += Environment.NewLine;
            message += "-----------------------------------------------------------";
            message += Environment.NewLine;
            string path ="C:/Unzip/EmailLog/ErrorLog.txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(message);
                writer.Close();
            }

        }
       protected void btnReset_Click(object sender, EventArgs e)
        {
            lblMessage.Text = lblExtract.Text = lblFailed.Text = lblTotal.Text = lblMail.Text = string.Empty;
            btnExtract.Enabled = true;
            btnReset.Visible = false;
        }
    

    }
}

    
  

