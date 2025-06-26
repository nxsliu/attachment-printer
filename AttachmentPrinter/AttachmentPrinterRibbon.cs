using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Spire.Pdf;
using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Printing;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace AttachmentPrinter
{
    public partial class AttachmentPrinterRibbon
    {
        private Outlook.Application outlookApplication; // Add a field to store the Outlook Application instance  

        private void AttachmentPrinterRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            outlookApplication = Globals.ThisAddIn.Application;
        }

        private void PrintButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (outlookApplication == null)
            {
                MessageBox.Show("Outlook Application is not initialized.");
                return;
            }

            MAPIFolder inbox = outlookApplication.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            Items unreadItems = inbox.Items.Restrict("[Unread]=true");

            MessageBox.Show(string.Format("Unread items in Inbox = {0}", unreadItems.Count));
            MailItem mailItem = null;

            int printedCount = 0;

            try
            {
                StringBuilder unsupportedFiles = new StringBuilder();
                foreach (object collectionItem in unreadItems)
                {
                    mailItem = collectionItem as MailItem;
                    if (mailItem != null && mailItem.Attachments.Count > 0)
                    {
                        foreach (Attachment attachment in mailItem.Attachments)
                        {
                            MessageBox.Show(string.Format("Attachment Name: {0}", attachment.FileName));

                            var extension = Path.GetExtension(attachment.FileName).ToLowerInvariant();

                            if (extension == ".pdf" || extension == ".docx" || extension == ".doc" || extension == ".jpeg" || extension == ".gif" || extension == ".png")
                            {
                                var tempFileName = mailItem.SentOn.ToString("s") + "-" + mailItem.SenderName + "-" + attachment.FileName;
                                tempFileName = tempFileName.Replace(":", "").Replace(" ", "");
                                var path = Directory.GetCurrentDirectory() + "\\AttachmentPrinter\\" + tempFileName;

                                if (File.Exists(path))
                                {
                                    File.Delete(path);
                                }
                                attachment.SaveAsFile(path);


                                //Print the document with the default printer 
                                if (extension == ".pdf")
                                {
                                    PrintPdf(path);
                                }
                                else if (extension == ".docx" || extension == ".doc")
                                {
                                    PrintWord(path);
                                }
                                else if (extension == ".jpeg" || extension == ".gif" || extension == ".png")
                                {
                                    PrintImage(path);
                                }

                                WaitForPrintQueueToClear();

                                printedCount++;
                                File.Delete(path);
                            }
                            else
                            {
                                unsupportedFiles.AppendLine(Environment.NewLine + $"Unsupported file type: {attachment.FileName}");
                            }
                        }
                        mailItem.UnRead = false;
                    }
                }

                MessageBox.Show($"🎉 Print process completed! Total attachments printed: {printedCount}" + unsupportedFiles.ToString());

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + Environment.NewLine + $"Attachments printed: {printedCount}");
            }
        }

        private void PrintPdf(string path)
        {
            PdfDocument doc = new PdfDocument(path);

            doc.Print();
        }

        private void PrintWord(string path)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Document doc = wordApp.Documents.Open(path);
            doc.PrintOut();
            doc.Close(false);
            wordApp.Quit(false);
        }

        private void PrintImage(string path)
        {
            PrintDocument pd = new PrintDocument();
            pd.PrintPage += (sender, args) =>
            {
                Image i = Image.FromFile(path);
                System.Drawing.Rectangle m = args.MarginBounds;

                if (i.Width / i.Height > m.Width / m.Height) // image is wider than page
                {
                    m.Height = i.Height / i.Width * m.Width;
                }
                else
                {
                    m.Width = i.Width / i.Height * m.Height;
                }
                args.Graphics.DrawImage(i, m);
            };
            pd.Print();
            pd.Dispose();
        }

        private void WaitForPrintQueueToClear()
        {
            using (LocalPrintServer printServer = new LocalPrintServer())
            {
                PrintQueue queue = LocalPrintServer.GetDefaultPrintQueue();
                queue.Refresh();

                const int maxWaitTimeMinute = 2;
                const int maxWaitTimeMs = maxWaitTimeMinute * 60000;
                int waited = 0;

                while (waited < maxWaitTimeMs)
                {
                    queue.Refresh();

                    if (queue.NumberOfJobs == 0 && !HasPrinterIssue(queue))
                        return;

                    if (HasPrinterIssue(queue))
                    {
                        throw new System.Exception("❌ Printer issue detected: " + GetPrinterErrorMessage(queue));
                    }

                    Thread.Sleep(1000);
                    waited += 1000;
                }

                throw new System.Exception($"⚠️ Print job did not complete after {maxWaitTimeMinute} minutes. Please check the printer and try again");
            }
        }

        private bool HasPrinterIssue(PrintQueue queue)
        {
            return queue.IsOffline || queue.IsOutOfPaper || queue.IsPaperJammed ||
                   queue.IsNotAvailable || queue.IsDoorOpened || queue.HasPaperProblem ||
                   queue.IsInError;
        }

        private string GetPrinterErrorMessage(PrintQueue queue)
        {
            if (queue.IsOffline) return "Printer is offline";
            if (queue.IsOutOfPaper) return "Out of paper";
            if (queue.IsPaperJammed) return "Paper jam";
            if (queue.HasPaperProblem) return "Paper problem";
            if (queue.IsDoorOpened) return "Printer door is open";
            if (queue.IsNotAvailable) return "Printer not available";
            if (queue.IsInError) return "Printer in error";

            return "Unknown printer issue";
        }
    }
}


