using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Spire.Pdf;
using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Printing;
using System.Reflection;
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

            LogInfo($"Scan started: unread items in Inbox = {unreadItems.Count}");
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
                            var extension = Path.GetExtension(attachment.FileName).ToLowerInvariant();

                            if (extension == ".pdf" || extension == ".docx" || extension == ".doc" || extension == ".jpeg" || extension == ".gif" || extension == ".png")
                            {
                                LogInfo($"Printing started: Sender name: {mailItem.SenderName}, Send on: {mailItem.SentOn.ToString("s")},  Attachment name: {attachment.FileName}");

                                var tempFileName = mailItem.SentOn.ToString("s") + "-" + mailItem.SenderName + "-" + attachment.FileName;
                                tempFileName = tempFileName.Replace(":", "").Replace(" ", "");
                                var path = Directory.GetCurrentDirectory() + "\\AttachmentPrinter\\" + tempFileName;

                                if (File.Exists(path))
                                {
                                    File.Delete(path);
                                }

                                LogInfo($"Save attachment to {path}");
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

                                LogInfo($"Printing complete: Sender name: {mailItem.SenderName}, Send on: {mailItem.SentOn.ToString("s")},  Attachment name: {attachment.FileName}");
                            }
                            else
                            {
                                LogInfo($"Unsupported file type: Sender name: {mailItem.SenderName}, Send on: {mailItem.SentOn.ToString("s")},  Attachment name: {attachment.FileName}");
                                unsupportedFiles.AppendLine(Environment.NewLine + $"Unsupported file type: {attachment.FileName}");
                            }
                        }
                        mailItem.UnRead = false;
                    }
                }

                LogInfo($"Scan completed: total attachments printed = {printedCount}");
                MessageBox.Show($"🎉 Print process completed! Total attachments printed: {printedCount}" + Environment.NewLine + unsupportedFiles.ToString());

            }
            catch (System.Exception ex)
            {
                LogException(ex);

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

        private void LogException(System.Exception ex)
        {
            try
            {
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Exception: {ex}\r\n";
                File.AppendAllText(GetLogFilePath(), logEntry);
            }
            catch
            {
                // Suppress any logging errors to avoid recursive exceptions
            }
        }

        private void LogInfo(string message)
        {
            try
            {
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Info: {message}\r\n";
                File.AppendAllText(GetLogFilePath(), logEntry);
            }
            catch
            {
                // Suppress any logging errors to avoid recursive exceptions
            }
        }

        private string GetLogFilePath()
        {
            string appDir = Directory.GetCurrentDirectory();
            return Path.Combine(appDir, $"Logs-{DateTime.Now:yyyy-MM-dd}.txt");
        }
    }
}


