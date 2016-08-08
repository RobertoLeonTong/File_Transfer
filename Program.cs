using System;
using IniParser;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using FtpLib;
using System.Windows.Forms;
//Roberto Leon Tong June 23rd 2016


namespace FileTransfer
{
    class Program
    {
        static void Main(string[] args)
        {
            Program n = new Program();
            FTP p = new FTP();
            try
            {
                //Initializes parser
                IniParser.FileIniDataParser parser = new FileIniDataParser();
                parser.CommentDelimiter = '#';
                IniData data = parser.LoadFile("config.ini");
                int Hours = Int32.Parse(data["Config"]["Hours"]);
                
                //Definitions
                Outlook.Application app = new Outlook.Application();
                Outlook.NameSpace ns = app.GetNamespace("MAPI");
                Outlook.MAPIFolder Inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                DateTime D1 = DateTime.Now.Subtract(new TimeSpan(0, 0, 0));
                DateTime D2 = DateTime.Now.Subtract(new TimeSpan(Hours, 0, 0));
                Outlook.Items ItemsA = Inbox.Items.Restrict("[ReceivedTime] >=" + D2.ToString("ddd M/dd/yyyy hh:mm tt") +'"');
                Outlook.Items Items2 = Inbox.Items;
                String[] Files = new string[10000];
                Boolean FilesFound = false;
                int i = 0;
                Boolean AllowSecondLoop = true;

                //Retrieves variables from config file
                string SaveTo = data["Config"]["SaveTo"];
                string Find1 = data["Config"]["Find1"];
                string Find2 = data["Config"]["Find2"];
                string Find3 = data["Config"]["Find3"];

                //Start of program
                Console.WriteLine("-------------------------------------------------");
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Program Start\n");
                Console.ResetColor();

                Console.WriteLine("\tFolder Name: {0}", Inbox.Name + "\n");
                Console.WriteLine("\tNumber of emails within " + Hours + " hours found: {0} \n\n\tFrom Time: {1}\n", ItemsA.Count, D2.ToString("yyyy-MM-dd hh:mm tt"));

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\nSubject lines to find\n");
                Console.ResetColor();

                Console.WriteLine("\t1) " + Find1 + "\n");
                Console.WriteLine("\t2) " + Find2 + "\n");
                Console.WriteLine("\t3) " + Find3 + "\n");

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\nFile(s) Found:\n");
                Console.ResetColor();

                //Regular expressions
                Regex regexA = new Regex(Find1, RegexOptions.IgnoreCase);
                Regex regexB = new Regex(Find2, RegexOptions.IgnoreCase);
                Regex regexC = new Regex(Find3, RegexOptions.IgnoreCase);

                //Searching all emails within inbox


                for (int counter = 1; counter <= ItemsA.Count ; counter++)
                { 
                    string x = Inbox.Items[counter].subject.Trim();
                    Match mA = regexA.Match(x);
                    Match mB = regexB.Match(x);
                    Match mC = regexC.Match(x);

                    //Regular expression to find current date in files
                    Regex regex1 = new Regex("(" + D1.ToString("yyyyMMdd") + ")");
                    
                    //Matching regular expressions
                    if (mA.Success || mB.Success || mC.Success )
                    {   
                        Outlook.MailItem mail = Inbox.Items[counter];
                        string x1 = mail.Attachments[1].FileName.Trim();
                        Match m1 = regex1.Match(x1);

                        if (m1.Success)
                        {
                            Files[i] = mail.Attachments[1].FileName;
                            
                            //Saves email attachment
                            if(Files[i] != null)
                                mail.Attachments[1].SaveAsFile(SaveTo + mail.Attachments[1].FileName);

                            FilesFound = true;
                            AllowSecondLoop = false;
                            i++;
                        }
                    }

               }

                for (int counter = Items2.Count; counter >= (Items2.Count - ItemsA.Count) && AllowSecondLoop == true; counter--)
                {
                    string x = Inbox.Items[counter].subject.Trim();
                    Match mA = regexA.Match(x);
                    Match mB = regexB.Match(x);
                    Match mC = regexC.Match(x);

                    //Regular expression to find current date in files
                    Regex regex1 = new Regex("(" + D1.ToString("yyyyMMdd") + ")");

                    //Matching regular expressions
                    if (mA.Success || mB.Success || mC.Success)
                    {
                        Outlook.MailItem mail = Inbox.Items[counter];
                        string x1 = mail.Attachments[1].FileName.Trim();
                        Match m1 = regex1.Match(x1);

                        if (m1.Success)
                        {
                            Files[i] = mail.Attachments[1].FileName;

                            //Saves email attachment
                            if (Files[i] != null)
                                mail.Attachments[1].SaveAsFile(SaveTo + mail.Attachments[1].FileName);

                            FilesFound = true;
                            i++;
                        }
                    }

                }

                //Displays all Files found
                foreach (string file in Files)
                    if(file != null)
                    {
                        Console.WriteLine("\t"+file.ToString());
                    }

                //Checks if files are found or not. If not it will not call the connection
                if (FilesFound == false) MessageBox.Show("No Files Found On " + DateTime.Now, "Notification");
                else
                {
                    p.loadconfig("config.ini");
                    p.Transferfiles(Files);
                }

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\nProgram Ended");
                Console.ResetColor();

                Console.WriteLine("-------------------------------------------------\n");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.StackTrace);
                MessageBox.Show("Process Failed!");
            }
        }

        
    }
}
