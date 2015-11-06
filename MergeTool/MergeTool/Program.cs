using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Drawing;
using System.Drawing.Imaging;
using BitMiracle.LibTiff.Classic;
using System.Runtime.CompilerServices;


namespace MergeTool
{
    class Program
    {
        public const String DocClass = "Late_3";
        private static String startTime;
        private static SqlConnection migrationDBCnn;
        private static String customErrMesg;
        private static Boolean threadResult;

        [STAThread]
        static void Main(string[] args)
        {
            startTime = DateTime.Now.ToString("yyyyMMddHHmmss");
            int totalDoc = 0;
            int successDoc = 0;

            //open a folder browser to select the destination root folder for all the merged tiff and pdf files
            FolderBrowserDialog fdb = new FolderBrowserDialog();
            fdb.Description = "Select the destination folder for exporting the merged files";
            fdb.SelectedPath = @"Y:\FairfaxStorage\Images\Exported Images";
            if (fdb.ShowDialog() != DialogResult.OK)
            {
                Environment.Exit(0);
            }
            string strDestinationRoot = fdb.SelectedPath;

            try
            {
                //connect to the MigrationDB and DestinationDB
                migrationDBCnn = new SqlConnection();
                Thread migrationDBThread = new Thread(new ThreadStart(() => initDBConnection("MigrationDB", "", "")));

                //reset thread flag
                threadResult = true;

                //start threads
                migrationDBThread.Start();
                migrationDBThread.Join();

                //check thread flag
                if (!threadResult)
                {
                    throw new Exception();
                }
                Console.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " => Database connected.");

                //get un-merged records from MigrationDB
                String sql = String.Format("select DocId, DOCS_Pages, FetchExportDirectory from is_docmap where MimeType = 'image/tiff' and FetchStatus = 'Success' and MergeStatus is null FetchExportDirectory is not null and DocClass = '{0}'", DocClass);
                DataRowCollection drSourceDocs = dbSelect(sql, migrationDBCnn).Tables[0].Rows;

                //prepare statistic variables
                totalDoc = drSourceDocs.Count;
                successDoc = 0;
                Console.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " => Found " + totalDoc + " docs ready to merge");

                //for merge tiff error handle
                Tiff.SetErrorHandler(new MyTiffErrorHandler());

                foreach (DataRow drSourceDoc in drSourceDocs)
                {
                    try
                    {
                        string strDocID = drSourceDoc[0].ToString();
                        string strPageNum = drSourceDoc[1].ToString();
                        string strSourceFolder = drSourceDoc[2].ToString();

                        Console.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " => Start merging doc " + strDocID + ".");

                        //check if the doc path exits
                        if (!Directory.Exists(strSourceFolder))
                        {
                            customErrMesg = String.Format("Cannot find folder {0} at {1}", strDocID, strSourceFolder);
                            logging(customErrMesg, "");
                            continue;
                        }

                        //get all the files under the path
                        List<string> lstImages = Directory.GetFiles(strSourceFolder).ToList();
                        lstImages.Sort();
                        Console.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " => Get all files for doc " + strDocID + ".");

                        //check if the page number match
                        if (!strPageNum.Equals(lstImages.Count.ToString()))
                        {
                            customErrMesg = String.Format("Doc {0} actual page number({1}) mismatch database record({2})", strDocID, lstImages.Count, strPageNum);
                            logging(customErrMesg, "");
                            continue;
                        }

                        //check if the folder is already in the destination root folder
                        //if it exits, delete it
                        string strDestinationFolder = strDestinationRoot + @"\" + strDocID;
                        if (Directory.Exists(strDestinationFolder))
                        {
                            customErrMesg = String.Format("Doc {0} already exits in destination root folder {1}, deleting it...", strDocID, strDestinationFolder);
                            logging(customErrMesg, "");

                            //delete all files under the directory
                            DirectoryInfo dirDestinationFolder = new DirectoryInfo(strDestinationFolder);
                            foreach (FileInfo file in dirDestinationFolder.GetFiles())
                            {
                                if (!file.Name.Equals("Thumbs.db"))
                                {
                                    file.Delete();
                                }
                            }
                        }
                        else
                        {
                            //create the destination folder and file name for the docs
                            Directory.CreateDirectory(strDestinationFolder);
                        }

                        string strDestinationFileNamePre = strDestinationFolder + @"\" + strDocID;

                        //start merging tiff and pdf
                        Thread tiff = new Thread(new ThreadStart(() => mergeTiff(lstImages, strDestinationFileNamePre)));
                        Thread pdf = new Thread(new ThreadStart(() => mergePdf(lstImages, strDestinationFileNamePre)));

                        //reset thread flag
                        threadResult = true;

                        //start threads
                        tiff.Start();
                        Console.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " => Start creating tiff file for doc " + ".");
                        pdf.Start();
                        Console.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " => Start creating pdf file for doc " + ".");
                        tiff.Join();
                        pdf.Join();

                        //check thread flag
                        if (!threadResult)
                        {
                            throw new Exception();
                        }

                        //update database records
                        sql = String.Format("update MigrationDB.dbo.is_docmap set MergeStatus = 'Success', MergeExportDirectory = '{0}' where DocId = {1}", strDestinationFolder, strDocID);
                        dbTransaction(sql, migrationDBCnn);
                        successDoc++;
                    }
                    catch (Exception e)
                    {
                        logging("Error occurred inside the loop", e.Message);
                        continue;
                    }
                }
            }
            catch (Exception e)
            {
                logging("Error occurred at main thread", e.Message);
            }
            finally
            {
                //close database connections
                migrationDBCnn.Close();
                migrationDBCnn.Dispose();
            }
            Console.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " => Merge process finished. Merged " + successDoc + " docs. Failed " + (totalDoc - successDoc) + " docs.");
            Console.ReadKey();
            Environment.Exit(0);
        }

        [MethodImpl(MethodImplOptions.Synchronized)]
        private static void MyThreadExcpetionHandler(String errMesg, Exception e)
        {
            logging(errMesg, e.Message);
            //update thread flag
            threadResult = false;
        }

        private static void logging(String customErrorMesg, string errorMesg)
        {
            string logPath = @"C:\log\" + startTime + ".txt";
            if (!File.Exists(logPath))
            {
                using (StreamWriter file = File.CreateText(logPath))
                {
                    file.WriteLine("tool started at " + startTime.ToString());
                }
            }
            using (StreamWriter file = File.AppendText(logPath))
            {
                string str = System.Environment.NewLine + DateTime.Now.ToString() + System.Environment.NewLine + customErrorMesg + System.Environment.NewLine + errorMesg;
                file.WriteLine(str);
                Console.WriteLine(str);
            }
        }

        private static void initDBConnection(string dbName, string dbUID, string dbPassword)
        {
            Console.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " => Connecting to database {0}", dbName);
            string strCnnString = String.Format("Data Source=USDXYSMRL1VW024;Initial Catalog={0};Integrated Security=True;", dbName, dbUID, dbPassword);

            try
            {
                migrationDBCnn.ConnectionString = strCnnString;
                migrationDBCnn.Open();
                migrationDBCnn.Close();
            }
            catch (Exception e)
            {
                customErrMesg = String.Format("Cannot establish database connection for {0}", dbName);
                MyThreadExcpetionHandler(customErrMesg, e);
            }
        }

        private static DataSet dbSelect(String sql, SqlConnection cnn)
        {
            DataSet ds = new DataSet();
            SqlDataAdapter da = new SqlDataAdapter();
            SqlCommand cmd = new SqlCommand();

            try
            {
                cnn.Open();
                cmd = new SqlCommand(sql, cnn);

                da.SelectCommand = cmd;
                da.Fill(ds);
            }
            catch (Exception e)
            {
                customErrMesg = String.Format(@"select issue for query: <{0}>", sql);
                logging(customErrMesg, e.Message);
                throw;
            }
            finally
            {
                cnn.Close();
            }

            return ds;
        }

        private static void dbTransaction(String sql, SqlConnection cnn)
        {
            cnn.Open();
            SqlCommand cmd = new SqlCommand();
            SqlTransaction trans = cnn.BeginTransaction();

            try
            {
                cmd.Connection = cnn;
                cmd.Transaction = trans;
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();

                trans.Commit();
            }
            catch (Exception e)
            {
                trans.Rollback();
                customErrMesg = String.Format(@"transaction issue for query: <{0}>", sql);
                logging(customErrMesg, e.Message);
                throw;
            }
            finally
            {
                cnn.Close();
            }
        }

        private static void mergePdf(List<string> lstImages, string strDestinationFileName)
        {
            strDestinationFileName += ".pdf";
            try
            {
                using (FileStream fs = new FileStream(strDestinationFileName, FileMode.Create, FileAccess.Write, FileShare.None))
                using (Document doc = new Document(PageSize.LETTER, 0, 0, 0, 0))
                {
                    PdfWriter writer = PdfWriter.GetInstance(doc, fs);
                    doc.Open();

                    //add the first page
                    iTextSharp.text.Image tiff = iTextSharp.text.Image.GetInstance(lstImages[0]);
                    tiff.ScaleToFit(doc.PageSize.Width, doc.PageSize.Height);
                    doc.Add(tiff);

                    //add the rest
                    for (int i = 1; i < lstImages.Count; i++)
                    {
                        doc.NewPage();
                        tiff = iTextSharp.text.Image.GetInstance(lstImages[i]);
                        tiff.ScaleToFit(doc.PageSize.Width, doc.PageSize.Height);
                        doc.Add(tiff);
                    }
                }

                //check the merged pdf page number
                PdfReader pdfReader = new PdfReader(strDestinationFileName);
                if (pdfReader.NumberOfPages != lstImages.Count)
                {
                    customErrMesg = String.Format("The page number of the merged pdf file {0} doesn't match database record", strDestinationFileName);
                    throw new PageNumMismatchException(customErrMesg);
                }
            }
            catch (PageNumMismatchException e)
            {
                MyThreadExcpetionHandler(e.Message, e);
            }
            catch (Exception e)
            {
                customErrMesg = String.Format("Pdf merging error occured at {0}", strDestinationFileName);
                MyThreadExcpetionHandler(customErrMesg, e);
            }
        }

        private static void mergeTiff(List<string> lstImages, string strDestinationFileName)
        {
            strDestinationFileName += ".tif";
            ImageCodecInfo codec = null;

            foreach (ImageCodecInfo cCodec in ImageCodecInfo.GetImageEncoders())
            {
                if (cCodec.CodecName == "Built-in TIFF Codec")
                {
                    codec = cCodec;
                    break;
                }
            }

            try
            {
                //creating the multi-page tiff
                using (EncoderParameters imagePararms = new EncoderParameters(1))
                {
                    imagePararms.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.MultiFrame);

                    using (System.Drawing.Image destinationImage = (System.Drawing.Image)(new Bitmap(lstImages[0])))
                    {
                        destinationImage.Save(strDestinationFileName, codec, imagePararms);
                        imagePararms.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.FrameDimensionPage);

                        for (int i = 1; i < lstImages.Count; i++)
                        {
                            using (System.Drawing.Image img = (System.Drawing.Image)(new Bitmap(lstImages[i])))
                            {
                                destinationImage.SaveAdd(img, imagePararms);
                            }
                        }

                        imagePararms.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.Flush);
                        destinationImage.SaveAdd(imagePararms);
                    }
                }

                //checking merged tiff file
                using (Tiff image = Tiff.Open(strDestinationFileName, "r"))
                {
                    //testing for corrupted tiff
                    int numberOfDirectories = image.NumberOfDirectories();
                    for (int i = 0; i < numberOfDirectories; ++i)
                    {
                        image.SetDirectory((short)i);

                        int width = image.GetField(TiffTag.IMAGEWIDTH)[0].ToInt();
                        int height = image.GetField(TiffTag.IMAGELENGTH)[0].ToInt();

                        int imageSize = height * width;
                        int[] raster = new int[imageSize];

                        if (!image.ReadRGBAImage(width, height, raster, true))
                        {
                            customErrMesg = String.Format("Corrupted tiff page found in merged tiff file at {0}", strDestinationFileName);
                            logging(customErrMesg, "");
                            throw new PageNumMismatchException();
                        }
                    }

                    //checking page number with database record
                    if (lstImages.Count != numberOfDirectories)
                    {
                        customErrMesg = String.Format("The page number of the merged tiff file {0} doesn't match database record", strDestinationFileName);
                        throw new PageNumMismatchException(customErrMesg);
                    }
                }
            }
            catch (PageNumMismatchException e)
            {
                MyThreadExcpetionHandler(e.Message, e);
            }
            catch (Exception e)
            {
                customErrMesg = String.Format("Tif merging error occured at {0}", strDestinationFileName);
                MyThreadExcpetionHandler(customErrMesg, e);
            }
        }
    }
}
