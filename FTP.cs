using System;
using System.IO;
using IniParser;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using FtpLib;
using System.Windows.Forms;

public class FTP
{
    static string Address;
    static string UserID;
    static string Password;
    static string ExecutablePath;
    static string SaveTo;
    static string DestDir;
    static string configFile;

    public void loadconfig(string Config)
    {
        try
        {
            IniParser.FileIniDataParser parser = new FileIniDataParser();
            parser.CommentDelimiter = '#';
            IniData data = parser.LoadFile(Config);

            //Retrieving variables from config file
                Address = data["Config"]["Address"];
                UserID = data["Config"]["UserID"];
                Password = data["Config"]["Password"];
                ExecutablePath = data["Config"]["ExecutablePath"];
                SaveTo = data["Config"]["SaveTo"];
                DestDir = data["Config"]["DestDir"];

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("\n\nHost Information\n");

            Console.ResetColor();
            Console.WriteLine("\tAddress: " + Address + "\n");
            Console.WriteLine("\tUserID: " + UserID + "\n");
            Console.WriteLine("\tSave to: " + SaveTo + "\n");
            Console.WriteLine("\tDestination directory: " + DestDir);

        }
        catch (Exception e)
        {
            Console.WriteLine("\n" + e.Message + "\n");
        }
    }

    public FTP(){}

    public void Transferfiles(string [] Files)
    {
        int FileCount = 0;
        int FilePassed = 0;

        try
        {   //Establishes connection to host
            using (FtpConnection ftp = new FtpConnection(Address, UserID, Password))
            {
                //Opens Connection
                ftp.Open();
                ftp.Login();

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\n\nTransfer Information\n");

                Console.ForegroundColor = ConsoleColor.DarkMagenta;
                //Checks if directory exists


                //FtpDirectoryInfo[] dir = ftp.GetDirectories();
                //Console.WriteLine(dir[0].Name);

                if (ftp.DirectoryExists(DestDir))
                {
                    Console.WriteLine("\tDirectory: Found");

                    //Sets current directory to defined directory
                    ftp.SetCurrentDirectory(DestDir);

                    for (int i = 0; i < Files.Count(); i++)
                    {
                        if (Files[i] == null) break;
                        FileCount++;
                    }

                    //files in current directory
                    FtpFileInfo[] dirFiles = ftp.GetFiles();
                    
                    Console.WriteLine("\n\tCurrent Files in Destination Directory: " + dirFiles.Length + "\n");

                    bool pass;

                    Console.WriteLine("\tFiles Within The Destination Directory:\n");

                    int x = 1;
                    if (dirFiles.Length > 0)
                    {
                        foreach (var file in dirFiles)
                        {
                            Console.WriteLine("\t " + " " + x + ") " + file.Name.Trim());
                            x++;
                        }
                    }
                    else
                    {
                        Console.WriteLine("\t No Files");
                    }
                    Console.WriteLine("\n");

                    //Stores the found files into defined directory
                    for (int i = 0; i < Files.Count(); i++)
                    {
                        pass = true;

                        if (Files[i] == null) break;

                        for(int y = 0; y < dirFiles.Length; y++)
                        {
                            Match passing = Regex.Match(dirFiles[y].Name.Trim(), Files[i],RegexOptions.IgnoreCase);
                            
                            if (passing.Success) {
                                pass = false;
                            }
                        }
                        
                        if(pass == true)
                        {
                            //ftp.PutFile(SaveTo + Files[i]);
                            FilePassed++;
                        }
                    }

                    Console.WriteLine("\tFiles Transferred: " + FilePassed + " file(s)\n");
                    MessageBox.Show("Notification:  " + FileCount + " file(s) have been found on " + DateTime.Now + "\nOnly " + FilePassed + " file(s) were passed.", "Notification");
                }
                else
                {
                    Console.WriteLine("\tDirectory Not Found");
                    MessageBox.Show("The Directory was not found.", "Notification");
                }

                //Closes connection
                ftp.Close();

                Console.ResetColor();//Resets Colours
            }
        }
        catch (Exception e)
        {
            Console.ResetColor();//Resets Colours
            Console.WriteLine("Message: " + e.StackTrace);

        }
    }

    //private bool DirectoryExists(string )
}
