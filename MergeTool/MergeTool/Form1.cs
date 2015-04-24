using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace MergeTool
{
    public partial class Form1 : Form
    {
        private SqlConnection migrationDBCnn;
        private SqlConnection destinationDBCnn;
        private DateTime startTime;

        public Form1()
        {
            InitializeComponent();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            destinationRootTB.Text = "";
            messageBoard.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fdb = new FolderBrowserDialog();

            fdb.Description = "Select the destination folder for exporting the merged files";
            fdb.SelectedPath = @"C:\Users\janetxue\Downloads\Migration\testing\DestinationTiff";

            if (fdb.ShowDialog() == DialogResult.OK)
            {
                destinationRootTB.Text = fdb.SelectedPath;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //clear message board, reset start time
            messageBoard.Text = "";
            startTime = DateTime.Now;

            if (string.IsNullOrEmpty(destinationRootTB.Text))
            {
                MessageBox.Show("Please select the destination root folder");
            }
            else
            {
                try
                {
                    migrationDBCnn = new SqlConnection();
                    destinationDBCnn = new SqlConnection();

                    //initial 2 database connections
                    initDbConnection("MigrationDB", "", "");
                    initDbConnection("DestinationDB", "", "");

                    //get un-merged records from MigrationDB
                    String sql = String.Format("select DocId, DOCS_Pages, FetchExportDirectory from is_docmap where MimeType = 'image/tiff' and FetchStatus = 'Success' and  MergeStatus is null and FetchExportDirectory is not null");
                    DataRowCollection drSourceDocs = dbSelect(sql, migrationDBCnn).Tables[0].Rows;

                    foreach (DataRow drSourceDoc in drSourceDocs)
                    {
                        try
                        {
                            string strDocId = drSourceDoc[0].ToString();
                            string strPageNum = drSourceDoc[1].ToString();
                            string strSourceFolder = drSourceDoc[2].ToString();

                            string test = "";
                            if (!Directory.Exists(strSourceFolder))
                            {
                                messageBoard.Text += String.Format("Cannot find folder {0} at {1}", strDocId, strSourceFolder);
                                logging("folder missing", strSourceFolder, "");
                            }
                            else
                            {
                                DirectoryInfo dirSourceFolder = new DirectoryInfo(strSourceFolder);
                                List<string> lstImages = Directory.GetFiles(strSourceFolder).ToList();

                                if (!strPageNum.Equals(lstImages.Count))
                                {
                                    messageBoard.Text += String.Format("Doc {0} doesn't match at {1}", strDocId, strSourceFolder);
                                    logging("doc number mis-match", strSourceFolder, "");
                                }
                                else
                                {
                                    string strDestinationFolder = destinationRootTB.Text + @"\" + strDocId;

                                    foreach (string image in lstImages)
                                    {
                                        test += image;
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                }
                catch (Exception ex)
                {
                }
                finally
                {
                    //close database connection
                    migrationDBCnn.Close();
                    migrationDBCnn.Dispose();
                    destinationDBCnn.Close();
                    destinationDBCnn.Dispose();
                }
            }
        }

        private void logging(string tiffOrPdf, string destination, string errorMesg)
        {
            string logName = startTime.ToString().Replace(" ", string.Empty).Replace("/", string.Empty).Replace(":", string.Empty);
            string logPath = @"C:\log\" + logName + ".txt";
            if (!File.Exists(logPath))
            {
                using (StreamWriter file = File.CreateText(logPath))
                {
                    file.WriteLine("tool started at " + startTime.ToString());
                }
            }
            using (StreamWriter file = File.AppendText(logPath))
            {
                file.WriteLine(tiffOrPdf + " error occured in folder " + destination);
                file.WriteLine(errorMesg);
                file.WriteLine();
            }
            throw new Exception();
        }

        private void initDbConnection(string dbName, string dbUID, string dbPassword)
        {
            string strCnnString = String.Format("Data Source=(local);Initial Catalog={0};Persist Security Info=True;User ID={1};Password={2};", dbName, dbUID, dbPassword);
            try
            {
                if (dbName.Equals("MigrationDB"))
                {
                    migrationDBCnn.ConnectionString = strCnnString;
                }
                else
                {
                    destinationDBCnn.ConnectionString = strCnnString;
                }
            }
            catch (Exception e)
            {
                messageBoard.Text += String.Format("Cannot initialize database {0}", dbName);
                logging("database connection issue", "", "check database " + dbName);
                throw e;
            }
        }

        private DataSet dbSelect(String sql, SqlConnection cnn)
        {
            DataSet ds = new DataSet();
            try
            {
                SqlDataAdapter da = new SqlDataAdapter();
                SqlCommand cmd;
                cmd = new SqlCommand(sql, cnn);

                da.SelectCommand = cmd;
                da.Fill(ds);
            }
            catch (Exception e)
            {
                logging("select issue", sql, "");
                throw e;
            }
            return ds;
        }

        private void mergePdf(string strDocId, List<string> lstImages, string strDestinationFolder, string strPageNum)
        {
            try
            {
                //create the dpf
                FileStream fs = new FileStream(strDestinationFolder, FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc = new Document(PageSize.LETTER, 0, 0, 0, 0);
                PdfWriter writer = PdfWriter.GetInstance(doc, fs);
                doc.Open();

                iTextSharp.text.Image tiff = iTextSharp.text.Image.GetInstance(lstImages[0]);
            }
            catch (Exception e)
            {
                messageBoard.Text += String.Format("Cannot merge pdf file for doc {0} at {1}", strDocId, strDestinationFolder);
                logging("pdf merge issue", strDestinationFolder, "failed to merge pdf");
                throw e;
            }
        }
    }
}
