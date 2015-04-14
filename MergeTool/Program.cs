﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System.IO;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace MergeTool
{
    class Program
    {
        static DateTime startTime;

        static void mergeTiffPages(string str_DestinationPath, List<string> sourceFiles)
        {
            str_DestinationPath += ".tif";
            System.Drawing.Imaging.ImageCodecInfo codec = null;

            foreach (System.Drawing.Imaging.ImageCodecInfo cCodec in System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders())
            {
                if (cCodec.CodecName == "Built-in TIFF Codec")
                    codec = cCodec;
            }

            try
            {
                System.Drawing.Imaging.EncoderParameters imagePararms = new System.Drawing.Imaging.EncoderParameters(1);
                imagePararms.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)System.Drawing.Imaging.EncoderValue.MultiFrame);

                if (sourceFiles.Count == 1)
                {
                    System.IO.File.Copy((string)sourceFiles[0], str_DestinationPath, true);

                }
                else if (sourceFiles.Count > 1)
                {
                    System.Drawing.Image DestinationImage = (System.Drawing.Image)(new System.Drawing.Bitmap((string)sourceFiles[0]));

                    DestinationImage.Save(str_DestinationPath, codec, imagePararms);

                    imagePararms.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)System.Drawing.Imaging.EncoderValue.FrameDimensionPage);


                    for (int i = 1; i < sourceFiles.Count; i++)
                    {
                        System.Drawing.Image img = (System.Drawing.Image)(new System.Drawing.Bitmap((string)sourceFiles[i]));

                        DestinationImage.SaveAdd(img, imagePararms);
                        img.Dispose();
                    }

                    imagePararms.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)System.Drawing.Imaging.EncoderValue.Flush);
                    DestinationImage.SaveAdd(imagePararms);
                    imagePararms.Dispose();
                    DestinationImage.Dispose();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("tiff error occured in folder " + str_DestinationPath);
                logging(str_DestinationPath, "tiff", e.ToString());
            }
        }

        static void mergePdf(string destination, List<string> sourceFiles)
        {
            try
            {
                destination += ".pdf";
                PdfDocument doc = new PdfDocument();
                XGraphics xgr;
                XImage img;
                for (int i = 0; i < sourceFiles.Count; i++)
                {
                    doc.AddPage();
                    xgr = XGraphics.FromPdfPage(doc.Pages[i]);
                    img = XImage.FromFile(sourceFiles[i]);
                    xgr.DrawImage(img, 0, 0);
                }
                doc.Save(destination);
                doc.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("pdf error occured in folder " + destination);
                logging(destination, "pdf", e.ToString());
            }
        }

        static void logging(string destination, string tiffOrPdf, string errorMesg)
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

        [STAThread]
        static void Main(string[] args)
        {
            startTime = DateTime.Now;

            string rootFolder;
            string destinationRoot;
            if (args.Length == 2)
            {
                rootFolder = args[0].Trim();
                destinationRoot = args[1].Trim();
            }
            else
            {
                FolderBrowserDialog fbd1 = new FolderBrowserDialog();
                FolderBrowserDialog fbd2 = new FolderBrowserDialog();
                fbd1.Description = "Select the root folder for all the fetched images";
                fbd1.SelectedPath = @"C:\Users\janetxue\Downloads\Migration\testing";
                fbd2.Description = "Select the destination root folder for exporting the merged files";
                fbd2.SelectedPath = @"C:\Users\janetxue\Downloads\Migration\testing";
                if (fbd1.ShowDialog() != DialogResult.OK)
                {
                    Environment.Exit(0);
                }
                if (fbd2.ShowDialog() != DialogResult.OK)
                {
                    Environment.Exit(0);
                }
                rootFolder = fbd1.SelectedPath;
                destinationRoot = fbd2.SelectedPath;
            }
            DirectoryInfo root = new DirectoryInfo(rootFolder);
            DirectoryInfo[] subdirs = root.GetDirectories();
            string[] images;
            List<string> docNames = new List<string>();
            foreach (DirectoryInfo subdir in subdirs)
            {
                docNames.Add(subdir.ToString());
            }
            docNames.Sort();
            foreach (string docName in docNames)
            {
                try
                {
                    //check if this doc is already merged
                    SqlConnection cnn = new SqlConnection();
                    cnn.ConnectionString = "Data Source=;Initial Catalog=;Integrated Security=True";
                    DataSet ds = new DataSet();
                    cnn.Open();
                    SqlDataAdapter da = new SqlDataAdapter();

                    String sql = String.Format("select FetchStatus from MigrationDB.dbo.is_docmap where DocId = {0}", docName);
                    SqlCommand cmd = new SqlCommand(sql, cnn);

                    da.SelectCommand = cmd;
                    da.Fill(ds);
                    cnn.Close();

                    DataTable dt = ds.Tables[0];
                    if (dt.Rows[0][0].ToString() != "done")
                    {
                        string destinationFolder = destinationRoot + @"\" + docName;
                        if (!Directory.Exists(destinationFolder))
                        {
                            Directory.CreateDirectory(destinationFolder);
                            images = Directory.GetFiles(rootFolder + @"\" + docName);
                            List<string> imagesList = images.ToList();
                            imagesList.Sort();
                            string destination = destinationFolder + @"\" + docName;

                            try
                            {
                                Thread tiff = new Thread(new ThreadStart(() => mergeTiffPages(destination, imagesList)));
                                tiff.Start();
                                Thread pdf = new Thread(new ThreadStart(() => mergePdf(destination, imagesList)));
                                pdf.Start();
                                tiff.Join();
                                pdf.Join();

                                //update the merged and converted information
                                sql = String.Format("update MigrationDB.dbo.is_docmap set MergeStatus = 'done', MergeExportDirectory = '{0}' where DocId = {1}", destinationFolder, docName);
                                SqlTransaction trans = cnn.BeginTransaction();

                                try
                                {
                                    cnn.Open();
                                    cmd = new SqlCommand();

                                    cmd.Connection = cnn;
                                    cmd.Transaction = trans;

                                    cmd.CommandText = sql;
                                    cmd.ExecuteNonQuery();

                                    trans.Commit();
                                    cnn.Close();
                                }
                                catch (SqlException e)
                                {
                                    logging(destinationFolder, "sql transaction error", "happened at docClass = " + docName);
                                    trans.Rollback();
                                    cnn.Close();
                                }
                            }
                            catch (Exception e)
                            {
                            }
                        }
                        else
                        {
                            logging(destinationFolder, "duplicate destination", "just delete the folder(s) in the destination folder");
                        }
                    }
                    cnn.Dispose();
                }
                catch (SqlException e)
                {
                    logging("no destination", "sql query/connection error", "happened at docClass = " + docName);
                }
            }
        }
    }
}
