using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Microsoft.WindowsAPICodePack.Dialogs;
using Outlook  =  Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;


namespace MailGenerator
{
    /// <summary>
    /// Logika interakcji dla klasy MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void buttonInvoiceSend_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // User's choice of file folder path
                CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                dialog.IsFolderPicker = true;
                string folderInvoicePath = "";

                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    folderInvoicePath = dialog.FileName; /*assigning the path to the folderInvoicePath variable*/
                }

                if (folderInvoicePath != "")
                {
                    string[] files = Directory.GetFiles(folderInvoicePath);

                    Microsoft.Office.Interop.Word._Application application = new Microsoft.Office.Interop.Word.Application() { Visible = false };

                    // Declaration of variables to mail
                    string mailEmployee = "";
                    string mailcustomer = "";
                    int countEmails = 0;
                    string invoiceName = "";

                    foreach (string file in files)
                    {
                        try
                        {
                            countEmails = 0;
                            Microsoft.Office.Interop.Word._Document document;

                            if (IsWordDocument(file)) /*file extension verification*/
                            {
                                object miss = Type.Missing;
                                object readOnly = true;
                                dynamic word = new Microsoft.Office.Interop.Word.Application();
                                object filename;

                                filename = file;
                                // Opening a Word file, read-only
                                document = application.Documents.Open(ref filename, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

                                // Verification of paragraphs in Word for the presence of customer / vendor email
                                foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in document.Paragraphs)
                                {
                                    if (paragraph.Range.Text.Contains("E-mail s") || paragraph.Range.Text.Contains("E-mail c"))
                                    {
                                        string[] words = paragraph.Range.Text.Split(' ');
                                        if (paragraph.Range.Text.Contains("E-mail s"))
                                        {
                                            mailEmployee = words[2];
                                            mailEmployee = mailEmployee.Substring(0, mailEmployee.Length - 2);
                                        }
                                        else
                                        {
                                            mailcustomer = words[2];
                                            mailcustomer = mailcustomer.Substring(0, mailcustomer.Length - 2);
                                        }
                                        countEmails++;
                                        // Stop the code when it finds both emails. 
                                        if (countEmails == 2)
                                        {
                                            break;
                                        }
                                    }
                                }

                                // Information about an erroneous file / missing data
                                if (countEmails != 2)
                                {
                                    MessageBox.Show("Incorrect file!");
                                }
                                else
                                {
                                    invoiceName = System.IO.Path.GetFileName(file);
                                    invoiceName = invoiceName.Substring(0, invoiceName.Length - 4);
                                    Outlook.Application app = new Outlook.Application(); /*creating a new outlook application*/
                                    Outlook.MailItem mail = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                                    mail.Subject = "Invoice " + invoiceName.ToString(); /*'Title'*/
                                    mail.To = mailcustomer;
                                    mail.CC = mailEmployee;

                                    // Email content encoded in HTML markup language
                                    mail.HTMLBody = @"<html><div style=""font-size:10.5px; font-family:Tahoma;"">" + "Dear Customer,<br>" + "" +
                                        "We are sending you an invoice with the number: <b>" + invoiceName.ToString() +
                                        "< The file is attached to this email." +
                                        "<p>Thank you for trusting our company. If you have any questions, please contact us at the following number: 000000000 or send an email to <a href=mailto:patrycja.krupa27@gmail.com>Customer helper</a>" +
                                        "<br>Check our website: <a href=https://github.com/pkrupa99>Visit Company</a>" +
                                        "<p>Best regards,<br>La Belle Flavour" + @"</div></html>";
                                    document.Close();
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(document);

                                    // Adding an invoice to an email
                                    string attachmentPath = filename.ToString();
                                    Outlook.Attachment attachment2 = mail.Attachments.Add(attachmentPath);
                                    mail.Display(true);

                                    // Moving an email to a folder 
                                    string moveFile = System.IO.Path.Combine(folderInvoicePath, "Send");
                                    MoveFileToFolder(file, moveFile);

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(attachment2);
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error processing file {file}: {ex.Message}");
                            // Log the exception
                        }
                    }

                    application.Quit();
                    MessageBox.Show("Finished");
                }
                else /*in the absence of selection, the message appears*/
                {
                    MessageBox.Show("You haven't chosen the path");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An unexpected error occurred: {ex.Message}");
                // Log the exception
            }
        }


        private void MoveFileToFolder(string filePath, string destinationFolder)
        {
            //create folder if doesn't exist
            if (!Directory.Exists(destinationFolder))
            {
                Directory.CreateDirectory(destinationFolder);
            }

            //download file name
            string fileName = System.IO.Path.GetFileName(filePath);

            //create path
            string destinationPath = System.IO.Path.Combine(destinationFolder, fileName);
            if (File.Exists(destinationPath))
            {
                MessageBox.Show($"File {fileName} already exists in the destination folder. Check it out and move the file");
                return;
            }
            try
            {
                File.Move(filePath, destinationPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error moving file: {ex.Message}");
            }
        }

        private bool IsWordDocument(string filePath)
        {
            string extension = System.IO.Path.GetExtension(filePath).ToLower();
            return extension == ".doc" || extension == ".docx";
        }

        private void buttonMailReport_Click(object sender, RoutedEventArgs e)
        {
            //declaration of excel application and worksheet
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets[1];

            //establishment of headings
            excelWorksheet.Cells[1, 1] = "Sender";
            excelWorksheet.Cells[1, 2] = "Title";
            excelWorksheet.Cells[1, 3] = "Date";

            //declaration first excel row 
            int excelRow = 2;


            //outlook initialization
            Outlook.Application oLk = new Outlook.Application();
            Outlook._NameSpace olNS = oLk.GetNamespace("MAPI");
            Outlook._Folders oFolders;
            oFolders = olNS.Folders;
            Outlook.MAPIFolder oFolder;
            oFolder = oFolders[1];
            Outlook.MAPIFolder oFolderIn = oFolder.Folders["Skrzynka odbiorcza"];
            Outlook.Items oItems = oFolderIn.Items;


            //loop to go through all the emails (unread)
            foreach (Outlook.MailItem mailItem in oFolderIn.Items.Restrict("[UnRead] = true"))
            {
                //Downloading e-mail information
                string senderMail = mailItem.SenderEmailAddress;
                string subject = mailItem.Subject;
                DateTime sentDate = mailItem.SentOn;

                //Saving data to an Excel spreadsheet
                excelWorksheet.Cells[excelRow, 1] = senderMail;
                excelWorksheet.Cells[excelRow, 2] = subject;
                excelWorksheet.Cells[excelRow, 3] = sentDate;

                excelRow++;
            }

            //Save Excel Workbook
            string pathExcel = $@"C:\Users\{Environment.UserName}\Downloads\Report_Excel";
            try
            {
                excelWorkbook.SaveAs(pathExcel);
                MessageBox.Show($"Report saved in location: {pathExcel}");
            }
            catch
            {
                MessageBox.Show("Please move file to another folder");
            }


            try
            {
                excelWorkbook.Close();
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error while closing Excel: {ex.Message}");
            }

            finally
            {
                //Excel Object Cleanup:
                Marshal.ReleaseComObject(excelWorksheet);
                Marshal.ReleaseComObject(excelWorkbook);
                Marshal.ReleaseComObject(excelApp);
            }
            

            
        }

        private void buttonMailRules_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Outlook application and components
                Outlook.Application oLk = new Outlook.Application();
                Outlook._NameSpace olNS = oLk.GetNamespace("MAPI");
                Outlook._Folders oFolders;
                oFolders = olNS.Folders;
                Outlook.MAPIFolder oFolder;
                oFolder = oFolders[1];

                Outlook.MAPIFolder oFolderIn = oFolder.Folders["Skrzynka odbiorcza"];
                Outlook.Items oItems = oFolderIn.Items;


                //rule statement
                string patternInvoice = @"\b(?:invoice|Inv\.|inv|vat)\b";
                string patternNewsletter = @"\b(?:newsletter|bulletin|digest|update|circular|mailer|briefing|report|dispatch)\b";
                string patternIssue = @"\b(?:issue|problem|challenge|concern|matter)\b";

                // Regular expressions
                Regex rgInvoice = new Regex(patternInvoice, RegexOptions.IgnoreCase);
                Regex rgNews = new Regex(patternNewsletter, RegexOptions.IgnoreCase);
                Regex rgIssue = new Regex(patternIssue, RegexOptions.IgnoreCase);

                //create folders in outlook if they do not exist
                Outlook.MAPIFolder oFolderInv = GetOrCreateFolder(oFolder, "Invoice");
                Outlook.MAPIFolder oFolderNews = GetOrCreateFolder(oFolder, "Newsletter");
                Outlook.MAPIFolder oFolderIssue = GetOrCreateFolder(oFolder, "Issue");

                //assignment of box contents
                var unreadMails = oFolderIn.Items.Restrict("[UnRead] = true").OfType<Outlook.MailItem>().ToList();

                //loop to verify emails unread by rules(sender / title)
                foreach (Outlook.MailItem mailItem in unreadMails)
                {
                    string senderMail = mailItem.SenderEmailAddress.ToLower();
                    string subject = mailItem.Subject.ToLower();


                    if (rgInvoice.IsMatch(senderMail) || rgInvoice.IsMatch(subject))
                    {
                        mailItem.Move(oFolderInv);

                    }
                    else if (rgNews.IsMatch(senderMail) || rgNews.IsMatch(subject))
                    {

                        mailItem.Move(oFolderNews);

                    }
                    else if (rgIssue.IsMatch(senderMail) || rgIssue.IsMatch(subject))
                    {
                        mailItem.Move(oFolderIssue);

                    }

                }
            }
            catch (Exception ex)
            {   
                MessageBox.Show($"Error while moving email: {ex.Message}");
            }

            MessageBox.Show("Emails sorted according to rules.Check your folders: Invoice, Newsletter, and Issue.");

        }

        private Outlook.MAPIFolder GetOrCreateFolder(Outlook.MAPIFolder parentFolder, string folderName)
        {
            Outlook.MAPIFolder subfolder;
            try
            {
                subfolder = parentFolder.Folders[folderName];
            }
            catch
            {
                subfolder = parentFolder.Folders.Add(folderName, Outlook.OlDefaultFolders.olFolderInbox);
            }
            return subfolder;
        }
    }
}
