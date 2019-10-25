using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using EAGetMail;
using GemBox.Spreadsheet;

namespace FunduszDomowy
{
    class Program
    {
        static Config oConfig = new Config();
        static IpkoConverter oIpkoConverter = new IpkoConverter();
        static Utf8Converter oUtf8Converter = new Utf8Converter();
        static GemBoxExcel oExcel = new GemBoxExcel();
        static string sBackupDir = string.Format("{0}\\inbox", Directory.GetCurrentDirectory()); //get inbox directory        

        static string _generateTemporaryFileName(int sequence, string from)
        {
            DateTime currentDateTime = DateTime.Now;
            return string.Format("{0}-{1:000}-{2:000}.txt",
                from,
                currentDateTime.ToString("yyyyMMddHHmmss", new CultureInfo("pl-PL")),
                sequence);
        }



        static void Main(string[] args)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            oExcel.CreateSpreadsheet();

            try
            {

                string sFullPath = "";
                string sLocalInbox = sBackupDir; // Create a folder named "inbox" under current directory
                                                 // to save the email retrieved.
                if (!Directory.Exists(sLocalInbox))
                { // If the folder is not existed, create it.
                    Directory.CreateDirectory(sLocalInbox);
                }

                MailServer oServer = new MailServer("imap.gmail.com", // Gmail IMAP4 server is "imap.gmail.com"
                                oConfig.Mail,
                                oConfig.Password,
                                ServerProtocol.Imap4);

                oServer.SSLConnection = true; // Enable SSL connection.
                oServer.Port = 993; // Set 993 SSL port

                MailClient oClient = new MailClient("TryIt");
                oClient.Connect(oServer);

                MailInfo[] infos = oClient.GetMailInfos();
                Console.WriteLine("Total {0} email(s)\r\n", infos.Length);
                for (int i = 0; i < infos.Length; i++)
                {

                    MailInfo info = infos[i];
                    Mail oMail = oClient.GetMail(info); // Receive email from IMAP4 server
                    DateTime dDate = oMail.ReceivedDate;

                    if (oMail.From.ToString().Contains(oConfig.Bank))
                    {
                        Console.WriteLine("Index: {0}; Size: {1}; UIDL: {2}",
                        info.Index, info.Size, info.UIDL);
                        Console.WriteLine("From: {0}", oMail.From.ToString());
                        Console.WriteLine("Subject: {0}\r\n", oMail.Subject);
                    }                        

                    string sMailContent = Regex.Replace(oMail.HtmlBody, @"\t|\n|\r", "");

                    if (oMail.From.ToString().Contains(oConfig.Bank))
                    { // if FROM contains our bank

                        string sFrom = Regex.Replace(oMail.From.ToString(), @"\s+", "").Replace("\"", "").Replace("<", "").Replace(">", "-"); // Add "from" to name of the mail
                        string sFileName = _generateTemporaryFileName(i + 1, sFrom); // Generate an unqiue email file name based on date time.
                        sFullPath = string.Format("{0}\\{1}", sLocalInbox, sFileName);

                        if (oMail.From.ToString().Contains(oConfig.Bank))
                            oMail.SaveAs(sFullPath, true); // Save email to local disk

                        using (StreamReader srFullMailTxt = new StreamReader(sFullPath))
                        {

                            string[] sLinesList = File.ReadAllLines(sFullPath);
                            bool bIsMatch = false;

                            for (int x = 0; x < sLinesList.Length - 1; x++)
                            {

                                if (oUtf8Converter.Utf8Convert(sMailContent) == sLinesList[x])
                                { //convert all to UTF-8, coz in mails are some polish chars
                                    srFullMailTxt.Close();
                                    oExcel.AddToSpreadsheet(oIpkoConverter.CutDate(dDate.ToString(), oIpkoConverter.CutMessage(sMailContent)));
                                    bIsMatch = true;
                                }
                            }

                            if (!bIsMatch)
                            {
                                srFullMailTxt.Close();
                                Console.WriteLine("nope, sry.."); // if value from match doesn't exist
                            }
                        }
                    }

                    if (oMail.From.ToString().Contains(oConfig.Bank))
                    { // function to delete values from gmail inbox after downloading them to local disc
                        oClient.Delete(info);
                    }
                }
                string[] sTxtList = Directory.GetFiles(sBackupDir, "*.txt");

                foreach (string f in sTxtList)
                { //delete all FULL maill txt files from directory
                    File.Delete(f);
                }

                // Quit and expunge emails marked as deleted from IMAP4 server.
                oClient.Quit();
                Console.WriteLine("Completed!");
            }
            catch (Exception ep)
            {
                Console.WriteLine(ep.Message);
                Console.ReadKey();
            }
        }
    }
}
