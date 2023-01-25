using System;
using System.IO;
using System.Net;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Linq;


namespace LabComm
{
    public class Message
    {
        public class ALab
        {
            public bool Request(string messagetype, XElement messagebody, out XDocument xmlresponce, out string errmessage)
            {
                bool success = true;
                xmlresponce = new XDocument();
                errmessage = "";
                Console.WriteLine("Составление запроса "); //Files.WriteLogFile("Составление XML для запроса");
                XDocument xdoc = new XDocument
                    (new XDeclaration("1.0", "windows-1251", ""),
                     new XElement("Message"
                    , new XAttribute("MessageType", messagetype)
                    , new XAttribute("Date", DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"))
                    , new XAttribute("Sender", Processing.Ini.mis)
                    , new XAttribute("Receiver", Processing.Ini.lis)
                    , new XAttribute("Password", Processing.Ini.WPass)
                                  , messagebody.Element("Version")
                                  , messagebody.Element("Query")
                                  , messagebody.Element("Patient")
                                  , messagebody.Element("Referral")
                                  , messagebody.Element("Assays")
                                  )
                     );

                try
                {
                    HttpWebRequest req = (HttpWebRequest)WebRequest.Create(Processing.Ini.url);
                    req.ContentType = "text/xml;charset=UTF-8";
                    req.Method = "POST";
                    Console.WriteLine("Отправка запроса"); //Files.WriteLogFile("Отправка запроса");
                    StreamWriter sw = new StreamWriter(req.GetRequestStream());
                    sw.WriteLine(xdoc);
                    sw.Close();
                    Console.WriteLine("Запрос отправлен, получение ответа"); //Files.WriteLogFile("Запрос отправлен, получение ответа");

                    WebResponse response = req.GetResponse();
                    xmlresponce = XDocument.Load(response.GetResponseStream());
 
                    //xmlresponce.Save("D:/data/responce.xml");
                    Console.WriteLine("Ответ получен"); //Files.WriteLogFile("Ответ получен");

                    if (xmlresponce.Element("Message").Attribute("MessageType").Value == "error")
                    {
                        success = false; Console.WriteLine("Ошибка входящего сообщения  {0}", xmlresponce.Element("Message"));
                    }
                }
                catch (WebException e)
                {
                    using (WebResponse response = e.Response)
                    {
                        HttpWebResponse httpResponse = (HttpWebResponse)response;
                        Console.WriteLine("Error code: {0}", httpResponse.StatusCode);
                        using (Stream data = response.GetResponseStream())
                        {
                            string text = new StreamReader(data).ReadToEnd();
                            Console.WriteLine(text);

                        }
                    }
                    success = false;
                }
                catch (NullReferenceException ex)
                {
                    Console.WriteLine("Processor Usage" + ex.Message);
                    success = false;
                }
                return success;
            }
        }

        public class SLS
        {
            public bool RequestGET(string messagebody, out XDocument xmlresponce, out string errmessage)
            {
                bool success = false;
                xmlresponce = null;
                errmessage = ""; Console.WriteLine("Отправка запроса...");
                string responsestr = string.Empty;
                try
                {
                    HttpWebRequest req = (HttpWebRequest)WebRequest.Create(messagebody);
                    req.ContentType = "text/xml;charset=windows-1251";
                    req.Method = "GET";
                    HttpWebResponse response = (HttpWebResponse)req.GetResponse();
 
                    Encoding Win1251 = Encoding.GetEncoding("windows-1251");

                    using (Stream stream = response.GetResponseStream())
                    {
                        using (StreamReader reader = new StreamReader(stream, Win1251))
                        {
                            char[] strresponse = reader.ReadToEnd().ToCharArray();
                            strresponse = strresponse.Where(x => x != 3 ).ToArray();
                            MemoryStream mStrm = new MemoryStream(Encoding.Default.GetBytes(strresponse));
                            responsestr = new string(strresponse);
                            xmlresponce = XDocument.Load(mStrm); //new XDocument();
                        }
                    }

                    response.Close(); success = true;
                }
                catch (WebException e)
                {
                    using (WebResponse response = e.Response)
                    {
                        HttpWebResponse httpResponse = (HttpWebResponse)response;
                        Console.WriteLine("Error code: {0}", httpResponse.StatusCode);
                        success = false;
                        using (Stream data = response.GetResponseStream())
                        {
                            errmessage = new StreamReader(data).ReadToEnd();
                            Console.WriteLine(errmessage);
                            Console.WriteLine(e.ToString());
                            Files.WriteLogFile(e.ToString());
                            Files.WriteLogFile(messagebody);
                        }
                    }
                }
                catch (Exception ex)
                {
                    success = false;
                    Console.WriteLine("Error: {0}", ex.Message);
                    Console.WriteLine(ex.ToString());
                    Files.WriteLogFile(ex.ToString());
                    Files.WriteLogFile(ex.Message);
                    Files.WriteLogFile(responsestr);

                }
                return success;
            }

            public bool RequestPost(XDocument xdoc, string filename, string Code, string Pass, out XDocument xmlresponce, out string errmessage)
            {
                bool success = true;
                xmlresponce = new XDocument();
                errmessage = "";
                try
                {
                    HttpWebRequest req = (HttpWebRequest)WebRequest.Create(string.Format("{0}/{1}", Processing.Ini.url, "upload_files.pl")); //string.Format("{0}/{1}", Processing.url, "upload_files.pl")
                    req.Method = "POST";
                    string boundary = DateTime.Now.Ticks.ToString("X");
                    req.ContentType = String.Format("multipart/form-data; boundary={0}", boundary);
                    req.KeepAlive = true;
                    req.UserAgent = "Mozilla Firefox 40.0";
                    req.Referer = "http://lablytech.ru/cgi-bin";
                    req.ProtocolVersion = HttpVersion.Version10;
                    ///*
                    string login = String.Format("--{0}\r\nContent-Type: {2}\r\nContent-Disposition: form-data; name=\"{1}\"\r\n\r\n{3}\r\n"
                        , boundary, "login", "text/plain; charset=utf-8", Code);
                    string wpass = String.Format("--{0}\r\nContent-Type: {2}\r\nContent-Disposition: form-data; name=\"{1}\"\r\n\r\n{3}\r\n"
                        , boundary, "password", "text/plain; charset=utf-8", Pass);
                    string fname = String.Format("--{0}\r\nContent-Type: {2}\r\nContent-Disposition: form-data; name=\"{1}\"\r\n\r\n{3}\r\n"
                        , boundary, "filename", "text/plain; charset=utf-8", filename);
                    //путь папки для заказа  [логин]/order
                    string target_dir = String.Format("--{0}\r\nContent-Type: {2}\r\nContent-Disposition: form-data; name=\"{1}\"\r\n\r\n{3}\r\n"
                        , boundary, "target_dir", "text/plain; charset=utf-8", "../obmen/"+ Code + "/order"); 
                    string content = String.Format("--{0}\r\nContent-Type: {2}\r\nContent-Disposition: form-data; name=\"{1}\"; filename=\"{3}\"\r\n\r\n"
                        , boundary, "file", "text/plain", filename);
                    string end = string.Format("\r\n--{0}", boundary);
                    
                    Console.WriteLine("Отправка запроса...."); //Files.WriteLogFile("Отправка запроса");
                    byte[] byteArray = Encoding.UTF8.GetBytes(login + wpass + fname + target_dir + content + xdoc.ToString() + end); //Encoding.UTF8.GetBytes("");
                    //req.ContentLength = byteArray.Length;

                    Stream sw = req.GetRequestStream();
                    sw.Write(byteArray, 0, byteArray.Length);
                    sw.Close();

                    Console.WriteLine("Запрос отправлен, получение ответа...."); //Files.WriteLogFile("Запрос отправлен, получение ответа");

                    WebResponse response = req.GetResponse();
                    xmlresponce = XDocument.Load(response.GetResponseStream());
                                     
                    Console.WriteLine("Ответ получен"); //Files.WriteLogFile("Ответ получен");

                }
                catch (WebException e)
                {
                    using (WebResponse response = e.Response)
                    {
                        HttpWebResponse httpResponse = (HttpWebResponse)response;
                        Console.WriteLine("Error code: {0}", httpResponse.StatusCode);
                        using (Stream data = response.GetResponseStream())
                        {
                            errmessage = new StreamReader(data).ReadToEnd();
                            Console.WriteLine(errmessage);
                         }
                    }
                    success = false;
                }
                catch (Exception ex)
                {
                    success = false;
                    Console.WriteLine("Error: {0}", ex.Message);
                    Files.WriteLogFile(ex.Message);
                }
                return success;

            }
        }

        public class Invitro
        {
            string request;
            string method;
            public Invitro(string request)
            {
                this.request = request;
            }

            public Invitro(string request, string method)
            {
                this.request = request;
                this.method = method;
            }

            public Invitro()
            { }
            public bool RegisterOrder(XDocument messagebody, out XDocument xmlresponce, out string errmessage)
            {
                bool success = true;
                xmlresponce = new XDocument();
                errmessage = "";
               
                try
                {
                    HttpWebRequest req = (HttpWebRequest)WebRequest.Create(Processing.Ini.url + "/RegisterOrder");
                    req.ContentType = "text/xml;charset=utf-8";
                    req.Method = "POST";
                    Console.WriteLine("Отправка запроса"); //Files.WriteLogFile("Отправка запроса");
                    StreamWriter sw = new StreamWriter(req.GetRequestStream());
                    sw.WriteLine(messagebody);
                    sw.Close();
                    Console.WriteLine("Запрос отправлен, получение ответа"); //Files.WriteLogFile("Запрос отправлен, получение ответа");

                    WebResponse response = req.GetResponse();
                    xmlresponce = XDocument.Load(response.GetResponseStream());

                    //xmlresponce.Save("D:/data/responce.xml");
                    Console.WriteLine("Ответ получен"); //Files.WriteLogFile("Ответ получен");
                }
                catch (WebException e)
                {
                    using (WebResponse response = e.Response)
                    {
                        HttpWebResponse httpResponse = (HttpWebResponse)response;
                        if (httpResponse != null)
                        {
                            Console.WriteLine("Error code: {0}", httpResponse.StatusCode);
                            using (Stream data = response.GetResponseStream())
                            {
                                errmessage = new StreamReader(data).ReadToEnd();
                                Console.WriteLine(errmessage);
                            }
                        }
                        else
                        {
                            Console.WriteLine("Error: {0}", "Пустой ответ сервера");
                            Files.WriteLogFile("Пустой ответ сервера");
                        }
                    }
                    success = false;
                }
                catch (Exception ex)
                {
                    success = false;
                    Console.WriteLine("Error: {0}", ex.Message);
                    Files.WriteLogFile(ex.Message);
                }
                return success;
            }

            public bool GetRequest(out XDocument xmlresponce, out string errmessage)
            {
                bool success = true;
                xmlresponce = new XDocument();
                errmessage = "";

                try
                {
                    HttpWebRequest req = (HttpWebRequest)WebRequest.Create(request);
                    req.ContentType = "text/xml;charset=utf-8";
                    req.Method = "Get";
                    Console.WriteLine("Отправка запроса"); 
                    Console.WriteLine("Запрос отправлен, получение ответа"); 

                    WebResponse response = req.GetResponse();
                    xmlresponce = XDocument.Load(response.GetResponseStream());

                    Console.WriteLine("Ответ получен"); 
                }
                catch (WebException e)
                {
                    /*using (WebResponse response = e.Response)
                    {
                        HttpWebResponse httpResponse = (HttpWebResponse)response;
                        Console.WriteLine("Error code: {0}", httpResponse.StatusCode);
                        using (Stream data = response.GetResponseStream())
                        {
                            errmessage = new StreamReader(data).ReadToEnd();
                            Console.WriteLine(errmessage);
                        }
                    }
                    success = false;*/
                }
                catch (Exception ex)
                {
                    success = false;
                    Console.WriteLine("Error: {0}", ex.Message);
                    Files.WriteLogFile(ex.Message);
                }
                return success;
            }

            public bool Request(out XDocument xmlresponce, out string errmessage)
            {
                bool success = true;
                xmlresponce = new XDocument();
                errmessage = "";

                try
                {
                    HttpWebRequest req = (HttpWebRequest)WebRequest.Create(request);
                    //req.ContentType = "text/xml;charset=utf-8";
                    req.Method = method;
                    req.KeepAlive = true;

                    Console.WriteLine("Отправка запроса");
                    Console.WriteLine("Запрос отправлен, получение ответа");

                    WebResponse response = req.GetResponse();
                    xmlresponce = XDocument.Load(response.GetResponseStream());

                    Console.WriteLine("Ответ получен");
                }
                catch (WebException e)
                {
                    using (WebResponse response = e.Response)
                    {
                        HttpWebResponse httpResponse = (HttpWebResponse)response;
                        if (httpResponse != null)
                        {
                            Console.WriteLine("Error code: {0}", httpResponse.StatusCode);
                            using (Stream data = response.GetResponseStream())
                            {
                                errmessage = new StreamReader(data).ReadToEnd();
                                Console.WriteLine(errmessage);
                            }
                        }
                        else
                        {
                            Console.WriteLine("Error: {0}", "Пустой ответ сервера");
                            Files.WriteLogFile("Пустой ответ сервера");
                        }
                    }
                    success = false;
                }
                catch (Exception ex)
                {
                    success = false;
                    Console.WriteLine("Error: {0}", ex.Message);
                    Files.WriteLogFile(ex.Message);
                }
                return success;
            }
        }

        public static string numberTo32sys(int ch)
        {
            switch (ch)
            {
                case 0: return "0";
                case 1: return "1";
                case 2: return "2";
                case 3: return "3";
                case 4: return "4";
                case 5: return "5";
                case 6: return "6";
                case 7: return "7";
                case 8: return "8";
                case 9: return "9";
                case 10: return "A";
                case 11: return "B";
                case 12: return "C";
                case 13: return "D";
                case 14: return "E";
                case 15: return "F";
                case 16: return "G";
                case 17: return "H";
                case 18: return "J";
                case 19: return "K";
                case 20: return "L";
                case 21: return "M";
                case 22: return "N";
                case 23: return "O";
                case 24: return "P";
                case 25: return "Q";
                case 26: return "R";
                case 27: return "S";
                case 28: return "T";
                case 29: return "U";
                case 30: return "V";
                case 31: return "W";
                default: return "";
            }
        }

        public static string GetMessageNumber()
        {
            DateTime now = DateTime.Now;
            string YY = now.Year.ToString().Substring(3);
            int dd = now.DayOfYear;

            return "0";
        }

        public static string DecToZex(int Dec, int  Base)
        {
            string Zex = string.Empty;
            int Rest;
            while (Dec > 0)
            {
                Rest = Dec % Base;
                Zex = (char)(Rest < 10 ? Rest + (int)('0') : Rest + (int)('A') - 10) + Zex;
                Dec = Dec / Base;
            }
            return Zex;
        }

    }
}
