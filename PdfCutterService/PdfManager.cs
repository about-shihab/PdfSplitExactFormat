using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.draw;
using iTextSharp.text.pdf.parser;
using Pechkin;
using Pechkin.Synchronized;
using Spire.Pdf;
using Spire.Pdf.AutomaticFields;
using Spire.Pdf.Graphics;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using static iTextSharp.text.pdf.PdfCopy;
using Font = System.Drawing.Font;

namespace PdfCutterService
{
    class PdfManager
    {
        public List<string> GetFileList(string type, string path)
        {
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            List<string> filePathList = Directory.GetFiles(path, type).ToList();
            return filePathList;
        }

        //pdf 
        public void ExtractPdf(string type)
        {
            try
            {
                string inputPdfPath = ConfigurationManager.AppSettings["pdfFolderPatrh"];
                string processedPdfPath = ConfigurationManager.AppSettings["processedFile"];
                string outputPdfPath = ConfigurationManager.AppSettings["outputPdfPath"];
                if (!Directory.Exists(outputPdfPath))
                    Directory.CreateDirectory(outputPdfPath);

                List<string> inputPdfList = this.GetFileList(type, inputPdfPath);
                List<string> processedPdfList = this.GetFileList(type, processedPdfPath);



                foreach (string pdfFile in inputPdfList)
                {

                    string inputFileName = pdfFile.Replace(inputPdfPath, processedPdfPath)+" Processed at "+DateTime.Now.ToString("dd MMMM yyyy HH.mm.ss")+".pdf";
                    if (processedPdfList.Contains(inputFileName))
                        continue;

                    string textContent = this.ReadPdfFile(pdfFile);



                    List<string> contentList = this.GetContentList(textContent, "Message Header", "End of Message");

                    foreach (string content in contentList)
                    {
                        string outputFileName = this.GetSubstring(content, "Documentary Credit Number", @"F31C:").Replace("<br>", "").Trim();
                        string outputFileFullPath = @outputPdfPath + "\\" + outputFileName + "@.pdf";
                        //ExportToPdf(FormatContent(headerText, content));
                        int indexOfFirstPage = content.IndexOf("Page") + "Page".Length;//index of page 
                        int indexOfFirstOf = content.IndexOf("of", indexOfFirstPage);// index of 'of' after page
                        int firstPageNum = Convert.ToInt32(content.Substring(indexOfFirstPage, indexOfFirstOf- indexOfFirstPage).Trim())-1;

                        int indexOfLastPage = content.LastIndexOf("Page") + "Page".Length;//index of page 
                        int indexOfLastOf = content.IndexOf("of", indexOfLastPage);// index of 'of' after page
                        int LastPageNum = Convert.ToInt32(content.Substring(indexOfLastPage, indexOfLastOf - indexOfLastPage).Trim());

                        this.SplitPdfByPage(pdfFile, outputFileFullPath, firstPageNum, LastPageNum);
                        this.AddPaging(outputFileFullPath);
                        //ExportToPdfByItext(this.FormatContent(content), outputFileFullPath);
                        this.WriteToFile(outputFileName + " \n file is splitted");
                    }
                    if (!Directory.Exists(processedPdfPath))
                        Directory.CreateDirectory(processedPdfPath);


                    //File.Move(pdfFile, inputFileName);
                    

                }
            }
            catch (Exception ex)
            {
                this.SendMail("MT700 Error Message:\n" + ex.Message, "MT700 Service Alert");
                throw ex;
            }

        }

        public string ReadPdfFile(string fileName)
        {
            StringBuilder text = new StringBuilder();

            if (File.Exists(fileName))
            {
                PdfReader pdfReader = new PdfReader(fileName);

                for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);

                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                    text.Append(currentText);
                }
                pdfReader.Close();
            }
            return text.ToString();
        }

        public List<int> AllIndexesOf(string str, string value)
        {
            if (String.IsNullOrEmpty(value))
                throw new ArgumentException("the string to find may not be empty", "value");
            List<int> indexes = new List<int>();
            for (int index = 0; ; index += value.Length)
            {
                index = str.IndexOf(value, index);
                if (index == -1)
                    return indexes;
                indexes.Add(index);
            }
        }


        public string GetSubstring(string content, string firstStr, string lastStr)
        {
            if (content.Contains(firstStr) && content.Contains(lastStr))
            {
                int firstIndex = content.IndexOf(firstStr) + firstStr.Length;
                int lastIndex = content.IndexOf(lastStr);
                return content.Substring(firstIndex, lastIndex - firstIndex);
            }
            else
            {
                return content;
            }
        }


        //create pdf by itextsharp

  


        private void SplitPdfByPage(string pdfFilePath, string outputPath, int startPage, int lastPage)
        {
            using (PdfReader reader = new PdfReader(pdfFilePath))
            {
                //iTextSharp.text.Rectangle rec = new Rectangle(100,100,100,100);
                Document document = new Document();
                PdfCopy copy = new PdfCopy(document, new FileStream(outputPath , FileMode.Create));

               // MemoryStream ms = new MemoryStream();

                document.Open();
                PdfImportedPage page;
                for (int pagenumber = startPage; pagenumber <= lastPage; pagenumber++)
                {
                    if (reader.NumberOfPages >= pagenumber)
                    {

                        page = copy.GetImportedPage(reader, pagenumber);


                        copy.AddPage(page);
                        //copy.AddPage(copy.GetImportedPage(reader, pagenumber));
                    }
                    else
                    {
                        break;
                    }

                }


                document.Close();

                
            }
        }



        public void AddPaging(string outputPath)
        {
            string destinationPath = outputPath.Replace(".pdf", "0.pdf");
            PdfReader reader = new PdfReader(outputPath);
            iTextSharp.text.Rectangle size = reader.GetPageSizeWithRotation(1);
            Document document = new Document(size);

            FileStream fs = new FileStream(destinationPath, FileMode.Create, FileAccess.Write);
            PdfWriter writer = PdfWriter.GetInstance(document, fs);
            document.Open();

            
            // create the new page and add it to the pdf
            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                document.NewPage();
                PdfContentByte cb = writer.DirectContent;
                PdfImportedPage page = writer.GetImportedPage(reader, i);
                cb.AddTemplate(page, 0, 0);
                cb.SetColorStroke(iTextSharp.text.BaseColor.GREEN);
                //cb.Rectangle(10, 50, page.Width-20, 30);
                cb.Stroke();
                cb.SetColorFill(BaseColor.WHITE);
                cb.Rectangle(10, 50, page.Width - 20, 30);

                cb.Fill();
                var baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.BeginText();
                cb.SetFontAndSize(baseFont, 9);
                cb.SetColorFill(BaseColor.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Page "+i+ " of "+ reader.NumberOfPages, 280, 20, 0);
                cb.EndText();
            }

            // close the streams and voilá the file should be changed :)
            document.Close();
            fs.Close();
            writer.Close();
            reader.Close();
            File.Delete(outputPath);
        }


  
        public List<string> GetContentList(string textContent, string firstString, string lastString)
        {
            List<int> firstIndex = this.AllIndexesOf(textContent, firstString);
            List<int> secondIndex = this.AllIndexesOf(textContent, lastString);

            List<string> contentList = new List<string>();

            if (firstIndex.Count > 0 && secondIndex.Count > 0)
            {
                for (int i = 0; i < firstIndex.Count; i++)
                {
                    string content = textContent.Substring(firstIndex[i], secondIndex[i] - firstIndex[i]);
                    contentList.Add(content);
                }
            }
            return contentList;
        }



        public void WriteToFile(string text)
        {
            string folderPath = ConfigurationManager.AppSettings["ServiceLog"].ToString();
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            string path = folderPath + "\\MT700_PdfCutter_ServiceLog.txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(string.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
                writer.Close();
            }
        }



        //for sending mail
        public void SendMail(string body, string subject)
        {
            try
            {
                SmtpClient client = new SmtpClient();
                client.Port = 25;
                client.Host = "hocs01.southeastbank.com.bd";
                client.EnableSsl = true;
                client.Timeout = 10000;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                MailMessage mm = new MailMessage("abdulla.mamun@southeastbank.com.bd", "abdulla.mamun@southeastbank.com.bd", subject, body);
                mm.CC.Add("abdulla.mamun@southeastbank.com.bd");
                mm.BodyEncoding = UTF8Encoding.UTF8;
                mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;

                client.Send(mm);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        

    }
}
