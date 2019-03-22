// This sample will open a PDF file that includes an image where the image is then extracted,
// sent to Project Oxford Vision API for OCR analysis.  The resulting text is then sent to an
// Azure Search index.

// Sample Code used based from the PDF Image extraction sample provided by 
// https://psycodedeveloper.wordpress.com/2013/01/10/how-to-extract-images-from-pdf-files-using-c-and-itextsharp/
// PDF Extraction done using iTextSharp
// These libraries may have different licenses then those for this sample

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license.
//
// Azure Search: http://azure.com
// Project Oxford: http://ProjectOxford.ai
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//

using AzureSearchOCR.AzureSearchOCR;
using Microsoft.Azure.Search;
using Microsoft.ProjectOxford.Vision.Contract;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using Xfinium.Pdf;
using Xfinium.Pdf.Rendering;
using Xfinium.Pdf.Rendering.RenderingSurfaces;
using Ghostscript.NET.Rasterizer;
using System;
using System.Drawing.Imaging;
using System.IO;
using System.Diagnostics;

namespace AzureSearchOCR
{
    class Program
    {
        /*
        static string oxfordSubscriptionKey = ConfigurationManager.AppSettings["oxfordSubscriptionKey"];
        static string searchServiceName = ConfigurationManager.AppSettings["searchServiceName"];
        static string searchServiceAPIKey = ConfigurationManager.AppSettings["searchServiceAPIKey"];
        static string indexName = "ocrtest";
        static SearchServiceClient serviceClient = new SearchServiceClient(searchServiceName, new SearchCredentials(searchServiceAPIKey));
        static SearchIndexClient indexClient = serviceClient.Indexes.GetClient(indexName) as SearchIndexClient;
        static VisionHelper vision = new VisionHelper(oxfordSubscriptionKey);
        */

        private static void CreateImageFromPage(SolidFramework.Pdf.Plumbing.PdfPage page, int dpi, string filename, int pageIndex, string extension, System.Drawing.Imaging.ImageFormat format)
        {
            // Create a bitmap from the page with set dpi.
            Bitmap bm = page.DrawBitmap(dpi);

            // Setup the filename.
            string filepath = string.Format(filename + "-{0}.{1}",
                                            pageIndex, extension);

            // If the file exits already, delete it. I.E. Overwrite it.
            if (File.Exists(filepath))
                File.Delete(filepath);

            // Save the file.
            bm.Save(filepath, format);

            // Cleanup.
            bm.Dispose();
        }
        private static void PdfToPng(string inputFile, string outputFileName, string outputFolder)
        {
            var xDpi = 300; //set the x DPI
            var yDpi = 300; //set the y DPI
            //var pageNumber = 1; // the pages in a PDF document

            using (var rasterizer = new GhostscriptRasterizer()) //create an instance for GhostscriptRasterizer
            {
                rasterizer.Open(inputFile); //opens the PDF file for rasterizing

                var pdfPages = rasterizer.PageCount;

                for (var i = 1; i <= pdfPages; i++)
                {
                    //set the output image(png's) complete path
                    var outputPNGPath = Path.Combine(outputFolder, string.Format("{0}-{1}.jpeg", outputFileName, i.ToString()));

                    //converts the PDF pages to png's 
                    var pdf2PNG = rasterizer.GetPage(xDpi, yDpi, i);

                    //save the png's
                    pdf2PNG.Save(outputPNGPath, ImageFormat.Jpeg);

                    Console.WriteLine("Saved " + outputPNGPath);
                }

            }
        }

        static void GetImagesFromPdf()
        {
            var searchPath = @"C:\Development\azure\OCR\AzureSearchOCR\AzureSearchOCR\AzureSearchOCR\pdf\";
            var outPath = @"C:\Development\azure\OCR\AzureSearchOCR\AzureSearchOCR\AzureSearchOCR\image\";
            var pdfFiles = Directory.GetFiles(searchPath, "*.pdf", SearchOption.TopDirectoryOnly);
            foreach (var pdfFile in pdfFiles)
            {
                string pdf_filename = searchPath + (pdfFile);

                Console.WriteLine("Extracting images from {0}", pdf_filename);

                var fileName = Path.GetFileNameWithoutExtension(pdfFile);

                PdfToPng(pdfFile, fileName, outPath);


                /*
                FileStream fs = new FileStream(pdf_filename, FileMode.Open, FileAccess.Read);
                PdfFixedDocument doc = new PdfFixedDocument(fs);
                PdfPageRenderer renderer = new PdfPageRenderer(doc.Pages[0]);
                PdfRendererSettings settings = new PdfRendererSettings();
                settings.DpiX = settings.DpiY = 96;
                PdfArgbIntRenderingSurface rs = renderer.CreateRenderingSurface<PdfArgbIntRenderingSurface>(settings.DpiX, settings.DpiY);
                settings.RenderingSurface = rs;
                renderer.ConvertPageToImage(settings);
                WriteableBitmap wb = new WriteableBitmap(rs.Width, rs.Height);
                Array.Copy(rs.Bitmap, wb.Pixels, rs.Bitmap.Length);
                */

                /*
                byte[] data = null;
                FileStream fs = new FileStream(pdf_filename, FileMode.Open, FileAccess.Read);
                BinaryReader br = new BinaryReader(fs);
                // Read the image into the byte variable
                data = br.ReadBytes((int)fs.Length);
                br.Close();
                fs.Close();
                */

                /*
                var images = PdfImageExtractor.ExtractImages(filename);
                Console.WriteLine("{0} images found.", images.Count);
                Console.WriteLine();
                var directory = System.IO.Path.GetDirectoryName(filename);
                foreach (var name in images.Keys)
                {
                    //if there is a filetype save the file
                    if (name.LastIndexOf(".") + 1 != name.Length)
                    {
                        var filePath = System.IO.Path.Combine(outPath, name);
                        images[name].Save(filePath);
                    }
                }
                */
            }
        }

        static void Main(string[] args)
        {
            /*
            var proc1 = new ProcessStartInfo();
            string Command;
            proc1.UseShellExecute = true;
            Command = @"ipconfig";
            proc1.WorkingDirectory = @"C:\Windows\System32";
            proc1.FileName = @"C:\Windows\System32\cmd.exe";
            /// as admin = proc1.Verb = "runas";
            proc1.Arguments = "/k " + Command;
            proc1.WindowStyle = ProcessWindowStyle.Maximized;
            Process.Start(proc1);
            */

            var pdfFilePath = @"C:\Development\azure\OCR\AzureSearchOCR\AzureSearchOCR\AzureSearchOCR\pdf\Thesis_Writing_Guidelines.pdf";
            var outFilePath = @"C:\Development\azure\OCR\AzureSearchOCR\AzureSearchOCR\AzureSearchOCR\image\Thesis_Writing_Guidelines-%d.png";
            var proc1 = new ProcessStartInfo();
            //string Command;
            proc1.UseShellExecute = true;
            //Command = @"ipconfig";
            proc1.WorkingDirectory = @"C:\Development\azure\OCR\AzureSearchOCR\AzureSearchOCR\AzureSearchOCR\image\";
            proc1.FileName = @"C:\Development\azure\OCR\pdf-to-jpg-master\bin\gswin32c.exe";
            var Arguments = "-dNOPAUSE -dBATCH -sDEVICE=png16m -sOutputFile=" + outFilePath + " -r300 -f " + pdfFilePath + " -c quit";

            Console.WriteLine("Arguments: ["+ Arguments + "]");
            /// as admin = proc1.Verb = "runas";
            proc1.Arguments = Arguments;
            proc1.WindowStyle = ProcessWindowStyle.Maximized;
            Process.Start(proc1);

            Console.WriteLine("All done.  Press any key to continue.");
            Console.ReadLine();

            /*
            GetImagesFromPdf();
            Console.WriteLine("Done");
            return;
            */


            /*
            var searchPath = "pdf";
            var outPath = "image";

            // Note, this will create a new Azure Search Index for the OCR text
            Console.WriteLine("Creating Azure Search index...");
            AzureSearch.CreateIndex(serviceClient, indexName);

            // Creating an image directory
            if (Directory.Exists(outPath) == false)
                Directory.CreateDirectory(outPath);

            foreach (var filename in Directory.GetFiles(searchPath, "*.pdf", SearchOption.TopDirectoryOnly))
            {
                Console.WriteLine("Extracting images from {0}", System.IO.Path.GetFileName(filename));
                var images = PdfImageExtractor.ExtractImages(filename);
                Console.WriteLine("{0} images found.", images.Count);
                Console.WriteLine();

                var directory = System.IO.Path.GetDirectoryName(filename);
                foreach (var name in images.Keys)
                {
                    //if there is a filetype save the file
                    if (name.LastIndexOf(".") + 1 != name.Length)
                        images[name].Save(System.IO.Path.Combine(outPath, name));
                }

                // Read in all the images and convert to text creating one big text string
                string ocrText = string.Empty;
                Console.WriteLine("Extracting text from image...");
                foreach (var imagefilename in Directory.GetFiles(outPath))
                {
                    OcrResults ocr = vision.RecognizeText(imagefilename);
                    ocrText += vision.GetRetrieveText(ocr);
                    File.Delete(imagefilename);
                }

                // Take the resulting orcText and upload to a new Azure Search Index
                // It is highly recommended that you upload documents in batches rather 
                // individually like is done here
                if (ocrText.Length > 0)
                {
                    Console.WriteLine("Uploading extracted text to Azure Search...");
                    string fileNameOnly = System.IO.Path.GetFileName(filename);
                    string fileId = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(fileNameOnly));
                    AzureSearch.UploadDocuments(indexClient, fileId, fileNameOnly, ocrText);
                }

            }

            // Execute a test search 
            Console.WriteLine("Execute Search...");
            AzureSearch.SearchDocuments(indexClient, "Azure Search");

            Console.WriteLine("All done.  Press any key to continue.");
            Console.ReadLine();
            */

        }
    }
}
