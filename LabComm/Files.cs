using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Xml.Linq;

namespace LabComm
{
    public class Files
    {
        static string FilePath;
        static public string CurrentDir;
        static public string LogsDir;
        static public string MessageDir;
        static public string OutputDir;
        static public string InputDir;
        static public string BadFilesDir;

        static public volatile Queue<string> QueueWriteToFile = new Queue<string>();
        static public volatile ManualResetEvent Write_Wait = new ManualResetEvent(false);

        public static void WriteLogFile(string Action)
        {
            DateTime CurrentDate = DateTime.Now;
            FilePath = Directory.GetCurrentDirectory() + "/logs" + "/log_" + CurrentDate.ToString("yyyy-MM-dd") + ".txt";
            StreamWriter file = new StreamWriter(FilePath, true);
            file.WriteLine(CurrentDate.ToString("HH:mm:ss") + ": " + Action);
            file.Close();
        }

        public static void RewriteIniFile(List<string> IniParameters, List<string> IniValues)
        {
            string line = null;

            List<string> IniLine = new List<string> { };

            FilePath = Directory.GetCurrentDirectory() + "/LabScan.ini";

            StreamReader rfile = new StreamReader(FilePath);
            while ((line = rfile.ReadLine()) != null) { IniLine.Add(line); }

            rfile.Close();
            for (int j = 0; j < IniParameters.Count; j++)
            {
                for (int i = 0; i < IniLine.Count; i++)
                {
                    if (IniLine[i].Contains(IniParameters[j])) IniLine[i] = IniParameters[j] + "=" + IniValues[j];
                }
            }

            File.Delete(FilePath);
            StreamWriter wfile = new StreamWriter(FilePath, true);

            foreach (string element in IniLine) wfile.WriteLine(element);
            wfile.Close();

        }

        public static void ReadIniFile(List<string> IniParameters, out List<string> IniValues)
        {
            string line = null;
            string value = null;
            bool newfile = false;

            //StringComparer StrCompare = StringComparer.OrdinalIgnoreCase;
            List<string> IniLine = new List<string> { };
            IniValues = new List<string> { };

            FilePath = Directory.GetCurrentDirectory() + "/LabScan.ini";

            StreamReader rfile = new StreamReader(FilePath);
            while ((line = rfile.ReadLine()) != null) { IniLine.Add(line); }

            rfile.Close();

            for (int j = 0; j < IniParameters.Count; j++)
            {
                for (int i = 0; i < IniLine.Count; i++)
                {
                    if (IniLine[i].Contains(IniParameters[j]))
                    {
                        value = IniLine[i].Substring(IniLine[i].IndexOf("=") + 1);
                        if (value == "")
                        {
                            newfile = true;
                            Console.Write("Input " + IniParameters[j] + ": ");
                            if (IniParameters[j] == "Password" || IniParameters[j] == "Web Pswd")
                            {
                                string str = string.Empty;
                                ConsoleKeyInfo key;
                                do
                                {
                                    key = Console.ReadKey(true);
                                    if (key.Key == ConsoleKey.Enter) break;
                                    if (key.Key == ConsoleKey.Backspace)
                                    {
                                        if (str.Length != 0)
                                        {
                                            str = str.Remove(str.Length - 1);
                                            Console.Write("\b \b");
                                        }
                                    }
                                    else
                                    {
                                        str += key.KeyChar;
                                        Console.Write("*");
                                    }
                                }
                                while (true);
                                value = str;
                                value = Crypto.EncryptStringAES(value, "lab");
                            }
                            else { value = Console.ReadLine(); }

                            IniLine[i] = IniParameters[j] + "=" + value;
                        }
                        if (IniParameters[j] == "Password" || IniParameters[j] == "Web Pswd") { value = Crypto.DecryptStringAES(value, "lab"); }

                    }
                }
                IniValues.Add(value); value = null;
            }

            if (newfile)
            {
                File.Delete(FilePath);
                StreamWriter wfile = new StreamWriter(FilePath, true);

                foreach (string element in IniLine)
                {
                    Console.WriteLine(element);
                    wfile.WriteLine(element);
                }
                wfile.Close();
            }
        }

        public static void WriteToBuffer(string RStr, bool Append)
        {
            FilePath = Directory.GetCurrentDirectory() + "/BufferR.txt";
            StreamWriter wfile = new StreamWriter(FilePath, Append);
            wfile.Write(RStr);
            wfile.Close();
        }

        public static void WriteToFile(string filename, string ToWrite, bool Append)
        {
            StreamWriter wfile = new StreamWriter(filename, Append);
            wfile.Write(ToWrite);
            wfile.Close();
        }

        public static void WriteToBuffer(XDocument RStr, bool Append)
        {
            FilePath = Directory.GetCurrentDirectory() + "/BufferR.txt";
            StreamWriter wfile = new StreamWriter(FilePath, Append);
            wfile.Write(RStr);
            wfile.Close();
        }

        public static void Hype()
        {
            string str = string.Empty;
            string hstr = "e1adc9f8668cedf99369c070ab2d9ca4";
            string csid;
            string xnd = "hip-hop-lab";
            MD5 mh = MD5.Create();
            if (!File.Exists("hype"))
            {
                ConsoleKeyInfo key;
                do
                {
                    key = Console.ReadKey(true);
                    if (key.Key == ConsoleKey.Enter) { Console.WriteLine(); break; }
                    if (key.Key == ConsoleKey.Backspace)
                    {
                        if (str.Length != 0)
                        {
                            str = str.Remove(str.Length - 1);
                            Console.Write("\b \b");
                        }
                    }
                    else
                    {
                        str += key.KeyChar;
                        Console.Write("*");
                    }
                }
                while (true);

                string h = GetH(mh, str);
                if (VerifyH(mh, str, hstr))
                {
                    csid = Crypto.EStr(Crypto.EncryptStringAES(DBControl.GetSID(), xnd));
                    StreamWriter file = new StreamWriter("hype");
                    file.WriteLine(csid);
                    File.SetAttributes("hype", FileAttributes.Hidden);
                    file.Close();
                }
                else
                {
                    Console.WriteLine("Неверное задание параметров запуска. Код ошибки: 0");
                    Thread.Sleep(6000);
                    Environment.Exit(1);
                }
            }
            else
            {
                string line;
                StreamReader rfile = new StreamReader("hype");
                line = rfile.ReadLine();

                if (Crypto.DecryptStringAES(Crypto.DStr(line), xnd) == DBControl.GetSID())
                {
                    Console.WriteLine("*********");
                }
                else
                {
                    Console.WriteLine("Неверное задание параметров запуска. Код ошибки: 1");
                    Thread.Sleep(6000);
                    Environment.Exit(1);
                }

                rfile.Close();
            }
        }

        public static bool GoFurther()
        {
            string License;
            string siddata = DBControl.GetSIDPlus();
            DateTime LicenseDate;
            CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
            bool result = true;
            try
            {
                if (!File.Exists("License"))
                {
                    WriteToFile("Code", Crypto.EStr(Crypto.EncryptStringAES(siddata, "hip-hop-lab")), false);
                    Files.WriteLogFile("Отсутствует лицензионный файл.");
                    Console.WriteLine("Отсутствует лицензионный файл.");
                    Thread.Sleep(6000); result = false;
                }
                else
                {
                    StreamReader rfile = new StreamReader("License");
                    License = rfile.ReadLine();
                    rfile.Close();
                    License = Crypto.DecryptStringAES(Crypto.DStr(License), "hip-hop-lab");
                    LicenseDate = DateTime.ParseExact(string.Format("20{0}", License.Substring(50, 6)), "yyyyMMdd", provider);

                    if (siddata != License.Substring(0, 47))
                    {
                        Files.WriteLogFile("Лицензионный ключ не подходит к коду базы, функционал приложения ограничен");
                        Console.WriteLine("Лицензионный ключ не подходит к коду базы, функционал приложения ограничен");
                        Thread.Sleep(6000); result = false;
                    }
                    if (Processing.Ini.DriverCode != License.Substring(47, 3))
                    {
                        Files.WriteLogFile("Лицензионный ключ не подходит к коду драйвера, функционал приложения ограничен");
                        Console.WriteLine("Лицензионный ключ не подходит к коду драйвера, функционал приложения ограничен");
                        Thread.Sleep(6000); result = false;
                    }
                    if (DateTime.Now > LicenseDate)
                    {
                        Files.WriteLogFile("Просроченный лицензионный ключ, функционал приложения ограничен");
                        Console.WriteLine("Просроченный лицензионный ключ, функционал приложения ограничен");
                        Thread.Sleep(6000); result = false;
                    }
                }
            }
            catch
            {
                Files.WriteLogFile("Неверная лицензия, запуск приложения невозможен");
                Console.WriteLine("Неверная лицензия, запуск приложения невозможен");
                Thread.Sleep(6000); result = false;
            }
            
            return result;
        }

        static string GetH(MD5 mh, string input)
        {
            byte[] data = mh.ComputeHash(Encoding.UTF8.GetBytes(input));
            StringBuilder sBuilder = new StringBuilder();
            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }
            return sBuilder.ToString();
        }

        public static string GetMD5(string input)
        {
            MD5 mh = MD5.Create();
            byte[] data = mh.ComputeHash(Encoding.UTF8.GetBytes(input));
            StringBuilder sBuilder = new StringBuilder();
            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }
            return sBuilder.ToString();
        }
        
        static bool VerifyH(MD5 mh, string input, string h)
        {
            string hashOfInput = GetH(mh, input);
            StringComparer comparer = StringComparer.OrdinalIgnoreCase;
            if (0 == comparer.Compare(hashOfInput, h)) { return true; }
            else { return false; }
        }

        public static List<string> FTPGetFileList(string path, string filepart, bool EnableSsl)
        {
            List<string> filelist = new List<string>();
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(Processing.Ini.url + path);
            request.EnableSsl = EnableSsl;

            request.Credentials = new NetworkCredential(Processing.Ini.mis, Processing.Ini.WPass);
            request.Method = WebRequestMethods.Ftp.ListDirectory;

            FtpWebResponse response = (FtpWebResponse)request.GetResponse();

            Stream responseStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream);
            string str;

            while (!reader.EndOfStream)
            {
                str = reader.ReadLine();
                if (str.IndexOf(filepart, StringComparison.OrdinalIgnoreCase) != -1)
                {
                    filelist.Add(str);
                }
            }
            
            reader.Close();
            responseStream.Close();
            response.Close();
            return filelist;
        }

        public static string FTPFolderInfo(string path, Boolean EnableSsl)
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(Processing.Ini.url + path);
            request.EnableSsl = EnableSsl;

            request.Credentials = new NetworkCredential(Processing.Ini.mis, Processing.Ini.WPass);
            request.Method = WebRequestMethods.Ftp.ListDirectory;

            FtpWebResponse response = (FtpWebResponse)request.GetResponse();

            Stream responseStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream);
            string str = reader.ReadToEnd();
            Console.WriteLine("Список фйалов для скачивания:");
            Console.WriteLine(str);
            return str; 
        }

        public static Stream FTPDownload(string path, bool EnableSsl)
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(Processing.Ini.url + path);
            request.EnableSsl = EnableSsl;
            request.Credentials = new NetworkCredential(Processing.Ini.mis, Processing.Ini.WPass);
            request.UsePassive = true;
            request.Method = WebRequestMethods.Ftp.DownloadFile;
            FtpWebResponse response = (FtpWebResponse)request.GetResponse();
            Stream dataStream = response.GetResponseStream();
            return dataStream;
        }

        public static XDocument FTPDownloadXML(string path, bool EnableSsl)
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(Processing.Ini.url + path);
            request.EnableSsl = EnableSsl;
            request.Credentials = new NetworkCredential(Processing.Ini.mis, Processing.Ini.WPass);
            request.UsePassive = true;
            request.Method = WebRequestMethods.Ftp.DownloadFile;
            FtpWebResponse response = (FtpWebResponse)request.GetResponse();
            Stream dataStream = response.GetResponseStream();
            XDocument xdoc = XDocument.Load(dataStream);
            dataStream.Close();
            response.Close();
            return xdoc;
        }

        public static void FTPDownloadFile(string ftppath, string filename, bool EnableSsl)
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(Processing.Ini.url + ftppath);
            request.EnableSsl = EnableSsl;
            request.Credentials = new NetworkCredential(Processing.Ini.mis, Processing.Ini.WPass);
            request.UsePassive = true;
            request.Method = WebRequestMethods.Ftp.DownloadFile;
            FtpWebResponse response = (FtpWebResponse)request.GetResponse();
            Stream dataStream = response.GetResponseStream();
            FileStream fileStream = File.OpenWrite(filename);
            dataStream.CopyTo(fileStream);
            dataStream.Close();
            fileStream.Close();
        }

        public static void FTPUpLoad(string path, Stream dataStream)
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(Processing.Ini.url + path);
            request.EnableSsl = true;
            request.Credentials = new NetworkCredential(Processing.Ini.mis, Processing.Ini.WPass);
            request.UsePassive = true;
            request.Method = WebRequestMethods.Ftp.UploadFile;
            Stream ftpStream = request.GetRequestStream();
            dataStream.CopyTo(ftpStream);
            ftpStream.Close();
        }

        public static void FTPDelete(string path, bool EnableSsl)
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(Processing.Ini.url + path);
            request.EnableSsl = EnableSsl;
            request.Credentials = new NetworkCredential(Processing.Ini.mis, Processing.Ini.WPass);
            request.UsePassive = true;
            request.Method = WebRequestMethods.Ftp.DeleteFile;

            FtpWebResponse response = (FtpWebResponse)request.GetResponse();
            Console.WriteLine("Delete status: {0}", response.StatusCode);
            response.Close();
        }

        public static void SaveStreamToFile(string filename, Stream stream)
        {
            FileStream fileStream = File.OpenWrite(filename);
            stream.CopyTo(fileStream);
            stream.Close();
            fileStream.Close();
        }

        public static Stream GetStreamFromFile(string filename)
        {
            FileStream fileStream = File.OpenRead(filename);
            return fileStream;
        }

        public static void WriteTraceFile()
        {
            while (Processing.proceed)
            {
                Write_Wait.Reset();
                while (QueueWriteToFile.Count != 0)
                {
                    string DataToWrite = QueueWriteToFile.Dequeue();
                    DateTime CurrentDate = DateTime.Now;
                    try
                    {
                        string fPath = MessageDir + "/buff_" + CurrentDate.ToString("yyyy-MM-dd") + ".txt";
                        StreamWriter wfile = new StreamWriter(fPath, true);
                        wfile.Write(DataToWrite);
                        wfile.Close();
                    }
                    catch (Exception ex)
                    {
                        WriteLogFile(ex.ToString()); Console.WriteLine(ex.Message);
                    }
                }
                Thread.Sleep(1);
                if (QueueWriteToFile.Count == 0) Write_Wait.WaitOne();
            }
        }

        public static void ToWriteFile(string s, UInt16 direction)
        {
            QueueWriteToFile.Enqueue(string.Format("{2}: {0}: {1} \r\n", DateTime.Now.ToString("HH:mm:ss.fff"), s, direction == 0 ? "Import": direction == 1 ? "Export" : "Message"));
            Write_Wait.Set();
        }
    }
}
