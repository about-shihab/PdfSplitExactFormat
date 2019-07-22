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
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
//using static iTextSharp.text.pdf.PdfCopy;
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

        public void ExtractDifferentPdf()
        {
            string type = "*.pdf";
            //MT700
            ExtractMT700Pdf(type, "MT700", "Message Header", "End of Message");
            ExtractMT202Pdf(type, "MT202", "Message Header", "Network delivery notif. request");
        }


        //MT700 pdf 
        public void ExtractMT700Pdf(string type,string projectNm,string firstMsg,string lastMsg)
        {
            try
            {

                string inputPdfPath = ConfigurationManager.AppSettings["input_pdfFolderPath"];
                string processedPdfPath = ConfigurationManager.AppSettings["processedFile"];
                string outputPdfPath = ConfigurationManager.AppSettings[projectNm + "_outputPdfPath"];
                if (!Directory.Exists(outputPdfPath))
                    Directory.CreateDirectory(outputPdfPath);

                List<string> inputPdfList = this.GetFileList(type, inputPdfPath);
                //List<string> processedPdfList = this.GetFileList(type, processedPdfPath).Select(s => s.Substring(0,s.IndexOf(" Processed at"))).ToList();



                foreach (string pdfFile in inputPdfList)
                {

                    string inputFileName = pdfFile.Replace(inputPdfPath, processedPdfPath);
                     //if (processedPdfList.Contains(inputFileName))
                     //   continue;

                    

                    string textContent = this.ReadPdfFile(pdfFile);
                    if (projectNm == "MT700" && !textContent.Contains("FIN 700"))
                        continue;

                    if (projectNm == "MT202" && !textContent.Contains("fin.202"))
                        continue;

                    List<string> contentList = this.GetContentList(textContent, firstMsg, lastMsg);

                    int unknownPageNumb = 1;
                    foreach (string content in contentList)
                    {
                        string outputFileName = this.GetSubstring(content, "Documentary Credit Number", @"F31C:").Replace("<br>", "").Trim();
                        
                        string outputFileFullPath = @outputPdfPath + "\\" + @outputFileName + "@0.pdf";
                        if (content.IndexOf("Page") != -1)
                        {
                            //find page number of Message Header
                            int indexOfFirstPage = content.IndexOf("Page") + "Page".Length;//index of page 
                            int indexOfFirstOf = content.IndexOf("of", indexOfFirstPage);// index of 'of' after page
                            int firstPageNum = Convert.ToInt32(content.Substring(indexOfFirstPage, indexOfFirstOf - indexOfFirstPage).Trim()) - 1;

                            //find page number of End of Message
                            int indexOfLastPage = content.LastIndexOf("Page") + "Page".Length;//index of page 
                            int indexOfLastOf = content.IndexOf("of", indexOfLastPage);// index of 'of' after page
                            int LastPageNum = Convert.ToInt32(content.Substring(indexOfLastPage, indexOfLastOf - indexOfLastPage).Trim());

                            this.SplitPdfByPageAddPaging(pdfFile, outputFileFullPath, firstPageNum, LastPageNum, projectNm);

                            unknownPageNumb = LastPageNum;
                        }
                        else
                        {
                            this.SplitPdfByPageAddPaging(pdfFile, outputFileFullPath, unknownPageNumb+1, unknownPageNumb+1, projectNm);
                        }
                        this.WriteToFile(outputFileName + " \n file is splitted");

                    }
                    if (!Directory.Exists(processedPdfPath))
                        Directory.CreateDirectory(processedPdfPath);


                    File.Move(pdfFile, inputFileName + projectNm+" Processed at " + DateTime.Now.ToString("dd MMMM yyyy HH.mm.ss") + ".pdf");


                }
            }
            catch (Exception ex)
            {
                this.SendMail("MT700 Error Message:\n" + ex.Message, "MT700 Service Alert");
                throw ex;
            }

        }


        public void ExtractMT202Pdf(string type, string projectNm, string firstMsg, string lastMsg)
        {
            try
            {

                string inputPdfPath = ConfigurationManager.AppSettings["input_pdfFolderPath"];
                string processedPdfPath = ConfigurationManager.AppSettings["processedFile"];
                string outputPdfPath = ConfigurationManager.AppSettings[projectNm + "_outputPdfPath"];
                if (!Directory.Exists(outputPdfPath))
                    Directory.CreateDirectory(outputPdfPath);

                List<string> inputPdfList = this.GetFileList(type, inputPdfPath);
                //List<string> processedPdfList = this.GetFileList(type, processedPdfPath).Select(s => s.Substring(0, s.IndexOf(" Processed at"))).ToList();



                foreach (string pdfFile in inputPdfList)
                {

                    string inputFileName = pdfFile.Replace(inputPdfPath, processedPdfPath);
                    //if (processedPdfList.Contains(inputFileName))
                    //    continue;

                    string textContent = this.ReadPdfFile(pdfFile);
                    if (projectNm == "MT202" && !textContent.Contains("fin.202"))
                        continue;

                    List<string> contentList = this.GetContentList(textContent, firstMsg, lastMsg);

                    int unknownPageNumb = 1;
                    foreach (string content in contentList)
                    {
                       
                        string lc_no = this.GetSubstring(content, "Transaction Reference:", "Related Reference:").Replace("<br>", "").Trim();
                        string bl_ref = this.GetSubstring(content, "Related Reference:", "Priority:").Replace("<br>", "").Trim();
                        string outputFileName = lc_no.Replace("/", "(slash)") + "@" + bl_ref.Replace("/", "(slash)");
                        string outputFileFullPath = @outputPdfPath + "\\" + @outputFileName + "@0.pdf";

                        if (content.IndexOf("Page") != -1)
                        {
                            //find page number of Message Header
                            int indexOfFirstPage = content.IndexOf("Page") + "Page".Length;//index of page 
                            int indexOfFirstOf = content.IndexOf("of", indexOfFirstPage);// index of 'of' after page
                            int firstPageNum = Convert.ToInt32(content.Substring(indexOfFirstPage, indexOfFirstOf - indexOfFirstPage).Trim()) - 1;

                            //find page number of End of Message
                            int indexOfLastPage = content.LastIndexOf("Page") + "Page".Length;//index of page 
                            int indexOfLastOf = content.IndexOf("of", indexOfLastPage);// index of 'of' after page
                            int LastPageNum = Convert.ToInt32(content.Substring(indexOfLastPage, indexOfLastOf - indexOfLastPage).Trim());

                            this.SplitPdfByPageAddPaging(pdfFile, outputFileFullPath, firstPageNum, LastPageNum, projectNm);

                            unknownPageNumb = LastPageNum;
                        }
                        else
                        {
                            this.SplitPdfByPageAddPaging(pdfFile, outputFileFullPath, unknownPageNumb + 1, unknownPageNumb + 1, projectNm);
                        }
                        this.WriteToFile(outputFileName + " \n file is splitted");

                    }
                    if (!Directory.Exists(processedPdfPath))
                        Directory.CreateDirectory(processedPdfPath);


                    File.Move(pdfFile, inputFileName+" _MT202 Processed at "+DateTime.Now.ToString("dd MMMM yyyy HH.mm.ss")+".pdf");


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



        public void SplitPdfByPageAddPaging(string pdfFilePath, string outputPath, int startPage, int lastPage, string projectNm)
        {
           // string destinationPath = outputPath.Replace(".pdf", "0.pdf");
            PdfReader reader = new PdfReader(pdfFilePath);
            iTextSharp.text.Rectangle size = reader.GetPageSizeWithRotation(1);
            Document document = new Document(size);

            FileStream fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
            PdfWriter writer = PdfWriter.GetInstance(document, fs);
            document.Open();

            int pageNumb = 1;
            // create the new page and add it to the pdf
            for (int i = startPage; i <= lastPage; i++)
            {
                document.NewPage();
                PdfContentByte cb = writer.DirectContent;
                PdfImportedPage page = writer.GetImportedPage(reader, i);
                cb.AddTemplate(page, 0, 0);
                cb.Stroke();
                cb.SetColorFill(BaseColor.WHITE);
                cb.Rectangle(10, 50, page.Width - 20, 30);
                cb.Fill();
                
                if(i == startPage && projectNm=="MT202")
                {
                    cb.SetColorFill(BaseColor.WHITE);
                    if (i==1)
                        cb.Rectangle(50, page.Height - 116, page.Width - 100, 55);
                    else
                       cb.Rectangle(50, page.Height - 88, page.Width - 100, 18);
                    cb.Fill();
                }
                

                var baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.BeginText();
                cb.SetFontAndSize(baseFont, 9);
                cb.SetColorFill(BaseColor.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Page "+ pageNumb + " of "+ (lastPage+1- startPage), 280, 20, 0);
                cb.EndText();
                pageNumb++;
            }

            // close the streams and voilá the file should be changed :)
            document.Close();
            fs.Close();
            writer.Close();
            reader.Close();
          //  File.Delete(outputPath);
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
                //client.Host = "hocs01.southeastbank.com.bd";
                client.Host = "webmail.southeastbank.com.bd";
                client.EnableSsl = true;
                client.Timeout = 10000;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential("tradewave@southeastbank.com.bd", "System1234");
                MailMessage mm = new MailMessage("tradewave@southeastbank.com.bd", "abdulla.mamun@southeastbank.com.bd", subject, body);
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
