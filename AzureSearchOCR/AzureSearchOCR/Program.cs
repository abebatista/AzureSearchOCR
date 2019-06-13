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
using System;
using System.Net.Http.Headers;
using System.Text;
using System.Net.Http;
using System.Web;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace AzureSearchOCR
{
    class Program
    {
        static string oxfordSubscriptionKey = ConfigurationManager.AppSettings["oxfordSubscriptionKey"];
        static string searchServiceName = ConfigurationManager.AppSettings["searchServiceName"];
        static string searchServiceAPIKey = ConfigurationManager.AppSettings["searchServiceAPIKey"];
        static string indexName = "ocrtest";
        static SearchServiceClient serviceClient = new SearchServiceClient(searchServiceName, new SearchCredentials(searchServiceAPIKey));
        static SearchIndexClient indexClient = serviceClient.Indexes.GetClient(indexName) as SearchIndexClient;
        static VisionHelper vision = new VisionHelper(oxfordSubscriptionKey);


        // Replace <Subscription Key> with your valid subscription key.
        const string subscriptionKey = "fb0555fe4eff4fddaa4ace6b1f7c8cbe";

        // You must use the same Azure region in your REST API method as you used to
        // get your subscription keys. For example, if you got your subscription keys
        // from the West US region, replace "westcentralus" in the URL
        // below with "westus".
        //
        // Free trial subscription keys are generated in the "westus" region.
        // If you use a free trial subscription key, you shouldn't need to change
        // this region.
        const string uriBase = "https://eastus2.api.cognitive.microsoft.com/vision/v2.0/ocr";

        /*

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

                
            }
        }

        static async Task<string> MakeRequest(string fileName)
        {
            var client = new HttpClient();

            // Request headers
            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "0859d36e0a604c2296438eac1fe430ab");

            // Request parameters
            var uri = "https://eastus.api.cognitive.microsoft.com/vision/v2.0/read/core/asyncBatchAnalyze?mode=Printed";

            HttpResponseMessage response;

            // Request body
            byte[] byteData = File.ReadAllBytes(fileName);

            string resp = "";
            using (var content = new ByteArrayContent(byteData))
            {
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                response = await client.PostAsync(uri, content);
                resp = await response.Content.ReadAsStringAsync();
            }
            return resp;
        }
        */



        const string pageOne = "https://bitnamieastus2symoyo.blob.core.windows.net/documentscontainer/1636959454949373196_fd4b3eaa-846e-4e5b-9149-bcc445a033d7.png";
        const string pageTwo = "https://bitnamieastus2symoyo.blob.core.windows.net/documentscontainer/1636959454975845655_cd64ad89-51e5-4213-81e4-ff25764263e9.png";
        /// <summary>
        /// Gets the text visible in the specified image file by using
        /// the Computer Vision REST API.
        /// </summary>
        /// <param name="imageFilePath">The image file with printed text.</param>
        static async Task MakeOCRRequest(string imageFilePath)
        {
            try
            {
                HttpClient client = new HttpClient();

                // Request headers.
                client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", subscriptionKey);

                // Request parameters. 
                // The language parameter doesn't specify a language, so the 
                // method detects it automatically.
                // The detectOrientation parameter is set to true, so the method detects and
                // and corrects text orientation before detecting text.
                //string requestParameters = "language=unk&detectOrientation=true";

                // Assemble the URI for the REST API method.
                string uri = "https://eastus.api.cognitive.microsoft.com/vision/v2.0/recognizeText?mode=Printed";

                HttpResponseMessage response;

                // Read the contents of the specified local image
                // into a byte array.
                byte[] byteData = GetImageAsByteArray(imageFilePath);

                // Add the byte array as an octet stream to the request body.
                using (ByteArrayContent content = new ByteArrayContent(byteData))
                {
                    // This example uses the "application/octet-stream" content type.
                    // The other content types you can use are "application/json"
                    // and "multipart/form-data".
                    content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

                    // Asynchronously call the REST API method.
                    response = await client.PostAsync(uri, content);
                }

                // Asynchronously get the JSON response.
                string contentString = await response.Content.ReadAsStringAsync();

                // Display the JSON response.
                Console.WriteLine("\nResponse:\n\n{0}\n",
                    JToken.Parse(contentString).ToString());
            }
            catch (Exception e)
            {
                Console.WriteLine("\n" + e.Message);
            }
        }

        static async Task GetOCROperationResult(string url)
        {
            try
            {
                HttpClient client = new HttpClient();

                // Request headers.
                client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", subscriptionKey);

                HttpResponseMessage response = await client.GetAsync(url);

                string contentString = await response.Content.ReadAsStringAsync();

                Console.WriteLine("\nResponse:\n\n{0}\n", JToken.Parse(contentString).ToString());

            }
            catch (Exception e)
            {
                Console.WriteLine("\n" + e.Message);
            }
        }

        /// <summary>
        /// Returns the contents of the specified file as a byte array.
        /// </summary>
        /// <param name="imageFilePath">The image file to read.</param>
        /// <returns>The byte array of the image data.</returns>
        static byte[] GetImageAsByteArray(string imageFilePath)
        {
            // Open a read-only file stream for the specified file.
            using (FileStream fileStream =
                new FileStream(imageFilePath, FileMode.Open, FileAccess.Read))
            {
                // Read the file's contents into a byte array.
                BinaryReader binaryReader = new BinaryReader(fileStream);
                return binaryReader.ReadBytes((int)fileStream.Length);
            }
        }

        static void Main(string[] args)
        {
            //GetOCROperationResult("https://eastus.api.cognitive.microsoft.com/vision/v2.0/textOperations/64ddb515-8fe9-490d-8cd2-a940389982e0");

            
            var currentDirectory = @"C:\TEMP\"; // AppDomain.CurrentDomain.BaseDirectory;
            var searchPath = currentDirectory; //  currentDirectory + "pdf";
            var outPath = currentDirectory; // currentDirectory + "image";

            foreach (var imagefilename in Directory.GetFiles(outPath, "*.png", SearchOption.TopDirectoryOnly))
            {

                //var imagefilenameString = MakeRequest(imagefilename).ConfigureAwait(true);                 
                //Console.WriteLine(imagefilenameString);
                //Get the path and filename to process from the user.

                Console.WriteLine("Optical Character Recognition:");
                Console.Write("Enter the path to an image with text you wish to read: ");
                var imageFilePath = imagefilename;


                if (File.Exists(imageFilePath))
                {
                    // Call the REST API method.
                    Console.WriteLine("\nWait a moment for the results to appear.\n");
                    MakeOCRRequest(imageFilePath).Wait();
                }
                else
                {
                    Console.WriteLine("\nInvalid file path");
                }


            }
            

            Console.WriteLine("All done.  Press any key to continue.");
            Console.ReadLine();


            /*
            foreach (var filename in Directory.GetFiles(searchPath, "*.pdf", SearchOption.TopDirectoryOnly))
            {

                // Read in all the images and convert to text creating one big text string
                string ocrText = string.Empty;
                Console.WriteLine("Extracting text from image...");
                foreach (var imagefilename in Directory.GetFiles(outPath, "*.png", SearchOption.TopDirectoryOnly))
                {
                    OcrResults ocr = vision.RecognizeText(imagefilename);
                    ocrText += vision.GetRetrieveText(ocr);
                    //File.Delete(imagefilename);
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
            */


            /*
            var pdfFilePath = @"C:\Development\pdf\Schedule1_387171.pdf";
            var outFilePath = @"C:\Development\pdf\Schedule1_387171-%d.png";
            var proc1 = new ProcessStartInfo();
            //string Command;
            proc1.UseShellExecute = true;
            //Command = @"ipconfig";
            proc1.WorkingDirectory = @"C:\Development\pdf\";
            proc1.FileName = @"C:\Development\azure\OCR\pdf-to-jpg-master\bin\gswin32c.exe";
            var Arguments = "-dNOPAUSE -dBATCH -sDEVICE=png16m -sOutputFile=" + outFilePath + " -r300 -f " + pdfFilePath + " -c quit";
            Console.WriteLine("Arguments: ["+ Arguments + "]");
            /// as admin = proc1.Verb = "runas";
            proc1.Arguments = Arguments;
            proc1.WindowStyle = ProcessWindowStyle.Maximized;
            Process.Start(proc1);
            Console.WriteLine("All done.  Press any key to continue.");
            Console.ReadLine();
            */

            /*
            var currentDirectory = AppDomain.CurrentDomain.BaseDirectory;

            var searchPath = currentDirectory + "pdf";
            var outPath = currentDirectory + "image";

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
