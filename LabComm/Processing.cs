using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
using System.Data.Linq;
using System.Linq;
using System.Xml.Linq;
using System.IO;
using System.Globalization;
using System.Timers;
using System.Text;
//using js = Newtonsoft.Json.Linq;
using Newtonsoft.Json;



namespace LabComm
{
    public class CheckInfo
    {
        public int OrderId;
        public int PatientId;
        public int FilialId;
    }

    public class Request
    {
        public int ID;
        public string Code;
        public int Nr;
        public DateTime Date;
        public int PatientID;
        public string FilialCode;
        public string FilialLabel;
        public string FilialPass;
        public string DepCode;
        public string DepLabel;
        public string DepId;
        public string WebId;
        public int Cito;
        public string Comment;
    }

    public class PatientInfo
    {
        public int ID;
        public string LastName;
        public string FirstName;
        public string MiddleName;
        public int Gender;
        public DateTime BirthDate;
        public int FazaCikla;
        public int SrokBeremennosti;
        public string workPlace;
        public string post;
        public string Email;
        public string phone;
        public string docTypeName;
        public int docTypeId = 1;
        public string docNumber;
        public string docSerNumber;
        public DateTime? issueDate;
        public string issueOrgName;
        public string snils;
        public string oms;
        public int adresType = 4;
        public string countryName;
        public int? countryId;
        public string regionName;
        public string regionType;
        public int? regionId;
        public string postalCode;
        public string district;
        public string districtType;
        public string town;
        public string townType;
        public int? cityId;
        public string streetName;
        public string streetType;
        public string house;
        public string building;
        public string appartament;
        public string addrReg;
        public string addrFact;
    }

    public class OrderInvitro
    {
        public int RequestId;
        public string RequestCode;
        public int? RequestNr;
        public DateTime RequestDate;
        public int FilialId;
        public string FilialCode;
        public string FilialLabel;
        public string ProductExternalId;
        public string BiomaterialOption;
        public string BiomaterialCode;
        public string AuxExtCode;
        public string AuxTableName;
        public string AuxFieldName;
        public int?  MotconsuId;
    }

    public class Processing
    {
        static public string sqlConnectionStr = string.Empty;
        static public int BarCodeCheckSum = 0;
        static public int DeleteCheckSum = 0;
        static public Thread WriteToFileThread;
 
        public class IniParameters
        {
            public int AnalyzerID;
            public string DriverCode;
            public string url;
            public string lis;
            public string mis;
            public string token;
            public string WPass;
            public int Version;

            public int OrgId;
            public int ModelId;
            public string BarcodeOption;
            public int ResGroupId;
            public int CultGroupId;
            public int AntbGroupId;
            public int BctrGroupId;
            public int RubricsId;
            public string CommentTable;
            public string CommentField;
            public int GetTimer;
            public int UploadTimer;
            public bool LoadDictionaries;
            public string ResultFilePath;
            public string OrderFilePath;
            public string ArchiveFilePath;
            public string PdfFilePath;
            public string DictFilePath;

            public bool WriteBuffer;
            public string HospitalID;
            public string HospitalName;
            public string HospitalCode;
            public bool DeleteDownloaded;
            public int TestPatientID;
            public int TestMode;
            public int Option;
            public bool WriteOutput;
            public bool WriteInput;
            public int OrderOption;
            public string UnboundPdfFilePath;
            public int MedecinId;
            public int PatdirRubricsId;
        }

        public static IniParameters Ini = new IniParameters();
        public static bool Upload = true;
        static public volatile bool proceed;

        static public int DepId;
        static public int MedDepId;
        static public int MedicinsId;
        static public string OrgCode;
        static public string Mesure;

 
        public static System.Timers.Timer OrderTimer;
        public static System.Timers.Timer ResultTimer;
        public static System.Timers.Timer bcChangeTimer;
        public static System.Timers.Timer delChangeTimer;

        public delegate void del();

        public class PatientBirthDate
        {
            public int BirthYear = 1900;
            public int BirthMonth = 1;
            public int BirthDay = 1;
        }

        public class PatientName
        {
            public string LastName;
            public string FirstName;
            public string MiddleName;
        }

        public class AssayInfo
        {
            public string Barcode;
            public string BiomaterialCode;
            public int BarcodeId;
            public string BiomaterialLabel;
            public int BioType;
            public string ProfileCode;
        }

        public class Order
        {
            public string Code;
            public int BarcodeId;
            public string Name;
            public int Obligatory;
            public string ExternalId;
            public int BioType;
            public string BioCode;
            public int GroupOption;
        }

        public class AddInfo
        {
            public string ProfileCode;
            public int GroupOption;
            public string InfoId;
            public string Value;
        }

        public class RequestOrder
        {
            public int RequestId;
            public string RequestCode;
            public int? RequestNr;
            public DateTime RequestDate;
            public int FilialId;
            public string FilialCode;
            public string FilialLabel;
            public int PatientID;
            public string LastName;
            public string FirstName;
            public string MiddleName;
            public int Gender;
            public DateTime BirthDate;
            public string ProductCode;
            public string ProductName;
            public int BarcodeId;
            public string BiomaterialOption;
            public string BiomaterialCode;
        }


        public class RestestData
        {
            public string VAL;
            public string UNIT;
            public string Rescode;
            public string ResName;
            public string NORM_TEXT_REC;
            public int State;
            public int Item;
            public string Comments;
            public int DataType;
            public string M_VAL;
            public string ResPage;
            public int MethodId;
        }

        public class Culture
        {
            public int Number;
            public string Label;
            public string Growth;
        }

        public class Antibiotic
        {
            public int Number;
            public string Label;
            public string SIR;
        }

        public class EMCTableData
        {
            public int MotconsuId;
            public int PatientId;
            public string TableName;
            public string FieldName;
            public string FieldValue;
        }

        public class WebFilialInfo
        {
            public string Code;
            public string WebId;
            public string Pass;
            public string Name;
        }

        public class ImpData
        {
            public int ImpdataId { get; set; }
            public string KEYCODE { get; set; }
            public string Nom { get; set; }
            public string Prenom { get; set; }
            public DateTime Date_Naissance { get; set; }
            public DateTime Date_Consultation { get; set; }
            public string Mesure { get; set; }
            public string FilialCode { get; set; }
            public string LabCode { get; set; }
        }

        public class Report
        {
            public string OrderId;
            public string Message;
        }

        public class Pages
        {
            public string SectionId;
            public string SectionName;
        }

        public class  AttachData
        {
            public int PatientId;
            public int MotconsuId;
            public DateTime ResDate;
        }

        public class AtachFile
        {
            public string Name;
            public string Body;
        }

        public static List<WebFilialInfo> FilialInfoList;

        public static del DLoadDictionaries;

        //Инициализация приложения
        public static void Initiation()
        {
            Files.CurrentDir = Directory.GetCurrentDirectory();
            Files.LogsDir = Files.CurrentDir + "/logs";
            if (!Directory.Exists(Files.LogsDir)) Directory.CreateDirectory(Files.LogsDir);
            Files.BadFilesDir = Files.CurrentDir + "/bad";
            if (!Directory.Exists(Files.BadFilesDir)) Directory.CreateDirectory(Files.BadFilesDir);
            sqlConnectionStr = DBControl.GetSqlConnectionString();

            Files.MessageDir = Files.CurrentDir + "/messages";
            if (Ini.WriteBuffer && !Directory.Exists(Files.MessageDir)) Directory.CreateDirectory(Files.MessageDir);
            Files.InputDir = Files.CurrentDir + "/input";
            if (Ini.WriteInput && !Directory.Exists(Files.InputDir)) Directory.CreateDirectory(Files.InputDir);
            Files.OutputDir = Files.CurrentDir + "/output";
            if (Ini.WriteOutput && !Directory.Exists(Files.OutputDir)) Directory.CreateDirectory(Files.OutputDir);

            //dbConnect = new SqlConnection(sqlConnectionStr);
            Files.WriteLogFile("Старт приложения");
            //Параметры инициализации
            List<string> Parameters = new List<string> { "Driver Code"      //0
                                                       , "Analyzer ID"      //1
                                                       , "LIS Url"
                                                       , "LIS Code"
                                                       , "MIS Code"         //4
                                                       , "Web Pswd"
                                                       , "Version"
                                                       , "Org Id"
                                                       , "Model Id"
                                                       , "ResGroup Id"
                                                       , "Get Timer"
                                                       , "Upload Timer"     //11
                                                       , "Result FilePath"  //12
                                                       , "Order FilePath"  
                                                       , "Write Buffer"     //14
                                                       , "Hospital ID"      //15
                                                       , "Archive FilePath" //16
                                                       , "Hospital Name"
                                                       , "Hospital Code"
                                                       , "Delete Downloaded"
                                                       , "Test Patient"
                                                       , "CultGroup Id"
                                                       , "AntbGroup Id"
                                                       , "Pdf FilePath"
                                                       , "Rubrics Id"
                                                       , "Comment Table"
                                                       , "Comment Field"
                                                       , "Barcode Option"
                                                       , "Load Dictionaries"  //28
                                                       , "Test Mode"
                                                       , "Dict FilePath"
                                                       , "Option"        //31
                                                       , "Write Input"        //32
                                                       , "Write Output"        //33
                                                       , "Order Option"        //34
                                                       , "Unbound FilePath"        //35
                                                       , "BctrGroup Id"
                                                       , "Medecin Id"              //37
                                                       , "Token"
                                                       , "PatdirRubrics Id"
            };
            List<string> Properties = null;
            //Задание параметров инициализации
            try
            {
                Files.ReadIniFile(Parameters, out Properties);
                Ini.DriverCode = Properties[0];
                Ini.AnalyzerID = Convert.ToInt32(Properties[1]);
                Ini.url = Properties[2];
                Ini.lis = Properties[3];
                Ini.mis = Properties[4];
                Ini.WPass = Properties[5];
                Ini.Version = Convert.ToInt32(Properties[6]);
                Ini.OrgId = Convert.ToInt32(Properties[7]);
                Ini.ModelId = Convert.ToInt32(Properties[8]);
                Ini.ResGroupId = Convert.ToInt32(Properties[9]);
                Ini.GetTimer = Convert.ToInt32(Properties[10]) * 60000;
                if (Char.IsNumber(Properties[11], 0))
                    Ini.UploadTimer = Convert.ToInt32(Properties[11]) * 60000;
                else Upload = false;
                Ini.ResultFilePath = Properties[12];
                Ini.OrderFilePath = Properties[13];
                Ini.WriteBuffer = (Properties[14] ?? "0") == "0" ? false : true;
                Ini.HospitalID = Properties[15];
                Ini.ArchiveFilePath = Properties[16];
                Ini.HospitalName = Properties[17];
                Ini.HospitalCode = Properties[18];
                Ini.DeleteDownloaded = (Properties[19] ?? "0") == "0" ? false : true;
                Ini.TestPatientID = (Properties[20] ?? "0") == "0" ? 0 : Convert.ToInt32(Properties[20]);
                Ini.CultGroupId =  Convert.ToInt32(Properties[21]);
                Ini.AntbGroupId =  Convert.ToInt32(Properties[22]);
                Ini.PdfFilePath = Properties[23];
                Ini.RubricsId = Convert.ToInt32(Properties[24]);
                Ini.CommentTable = Properties[25];
                Ini.CommentField = Properties[26];
                Ini.BarcodeOption = Properties[27] ?? "0";
                Ini.LoadDictionaries = (Properties[28] ?? "0") == "0" ? false : true;
                Ini.TestMode = (Properties[29] ?? "0") == "0" ? 0 : Convert.ToInt32(Properties[29]);
                Ini.DictFilePath = Properties[30];
                Ini.Option = (Properties[31] ?? "0") == "0" ? 0 : Convert.ToInt32(Properties[31]);
                Ini.WriteInput = (Properties[32] ?? "0") == "0" ? false : true;
                Ini.WriteOutput = (Properties[33] ?? "0") == "0" ? false : true;
                Ini.OrderOption = (Properties[34] ?? "0") == "0" ? 0 : Convert.ToInt32(Properties[34]);
                Ini.UnboundPdfFilePath = Properties[35];
                Ini.BctrGroupId = Convert.ToInt32(Properties[36]);
                Ini.MedecinId = (Properties[37] ?? "0") == "0" ? 0 : Convert.ToInt32(Properties[37]);
                Ini.token = Properties[38];
                Ini.PatdirRubricsId = (Properties[39] ?? "0") == "0" ? 0 : Convert.ToInt32(Properties[39]);
            }
            catch(Exception ex)
            {
                Files.WriteLogFile("Неверное задание параметров инициализации: "+ ex.ToString());
                Console.WriteLine("Неверное задание параметров инициализации");
                Console.WriteLine(ex.Message);
            }

            proceed = true;
            
            //Определение пользователя, отделения и группы методик
            try
            {
                DataContext db = new DataContext(sqlConnectionStr);
                //Table<Dep> TDep = db.GetTable<Dep>();
                //Table<MedDep> TMedDep = db.GetTable<MedDep>();
                Table<LAB_DEVS> TLAB_DEVS = db.GetTable<LAB_DEVS>();


                if (DBControl.GetDepId(out int medicinsId, out int depId, out int medDepId))
                {
                    MedDepId = medDepId;
                    DepId = depId;
                    MedicinsId = medicinsId == 0 ? Ini.MedecinId : medicinsId;
                }
                else
                {
                    Console.WriteLine("Должны использоваться настройки филиалов!");
                    Files.WriteLogFile("Должны использоваться настройки филиалов!");
                }

                /*
                var qDep = (from de in TDep
                            join me in TMedDep on de.DepId equals me.DepId
                            where de.OrgId == Ini.OrgId
                            select new
                            {
                                DepId = me.DepId,
                                MedDepId = me.MedDepId,
                                MedecinsId = me.MedecinsId
                            }).FirstOrDefault();

                if (qDep == null)
                {
                    Console.WriteLine("Неверное задание идентификаторов организации!");
                    Files.WriteLogFile("Неверное задание идентификаторов организации!");
                    //Thread.Sleep(5000);
                    //Environment.Exit(1);
                }

                MedDepId = qDep.MedDepId;
                DepId = qDep.DepId;
                MedicinsId = qDep.MedecinsId==0?Ini.MedecinId: qDep.MedecinsId;
                */

                OrgCode = db.ExecuteQuery<string>(@"
                select CODE from FM_ORG where FM_ORG_ID = {0}
                            ", Ini.OrgId).FirstOrDefault();

                if (OrgCode == null)
                {
                    Console.WriteLine("Неверное задание идентификаторов внешней лаборатории!");
                    Files.WriteLogFile("Неверное задание идентификаторов внешней лаборатории!");
                    Thread.Sleep(5000);
                    Environment.Exit(1);
                }

                Mesure = (from ge in TLAB_DEVS
                                 where ge.LabDevsId == Ini.AnalyzerID
                                 select ge.Label ?? "").SingleOrDefault();

                Console.WriteLine("Код организации: {0}    Группа методик: {1}  Код драйвера: {2} ", OrgCode, Mesure, Ini.DriverCode);
            }
            catch (SqlException ex)
            {
                Console.WriteLine("Проблемы с подключением к серверу!");
                Files.WriteLogFile(ex.ToString());
                Thread.Sleep(5000);
                Environment.Exit(1);
            }
            catch (Exception ex)
            {
                Files.WriteLogFile(ex.ToString()); Console.WriteLine(ex.Message);
            };
 
            del DLoadResults = new del(NULL_Action);
            del DUploadOrders = new del(NULL_Action);
            DLoadDictionaries = new del(NULL_DictAction);

            if (Ini.DriverCode == "ALab")
            {
                DLoadResults = new del(ALab.LoadResults);
                DUploadOrders = new del(ALab.UploadOrders);
            }
            if (Ini.DriverCode == "CLD")
            {
                DLoadResults = new del(CLD.LoadResults);
                DUploadOrders = new del(CLD.UploadOrders);
                //DLoadResults = new del(CLD.LoadResultsJson);
                //DUploadOrders = new del(CLD.UploadOrdersJson);
            }
            if (Ini.DriverCode == "SLS")
            {
                Files.MessageDir = Files.CurrentDir + "/messages";
                if (Ini.WriteBuffer && !Directory.Exists(Files.MessageDir)) Directory.CreateDirectory(Files.MessageDir);
                Files.InputDir = Files.CurrentDir + "/input";
                if (Ini.WriteInput && !Directory.Exists(Files.InputDir)) Directory.CreateDirectory(Files.InputDir);
                Files.OutputDir = Files.CurrentDir + "/output";
                if (Ini.WriteOutput && !Directory.Exists(Files.OutputDir)) Directory.CreateDirectory(Files.OutputDir);
                //Получение списка филиалов с паролями доступа для вэб сервиса
                FilialInfoList = DBControl.GetWebFilealsInfo();
                //Загрузка результатов и выгрузка заказов
                DLoadResults = new del(SLS.LoadResults);
                DUploadOrders = new del(SLS.UploadOrders);
                if (Ini.WriteBuffer)
                {
                    WriteToFileThread = new Thread(Files.WriteTraceFile);
                    WriteToFileThread.Start();
                }
            }
            if (Ini.DriverCode == "INV")
            {
                DLoadDictionaries = new del(Invitro.LoadDictionaries);
                DLoadResults = new del(INVITRO.LoadResults);
                //DUploadOrders = new del(Invitro.UploadOrders);
                //DUploadOrders = new del(Invitro.GetLabelInfo);

                bool delch = DBControl.delChanged(out int delChecksum);
                Processing.DeleteCheckSum = delChecksum;
                SetbcChangeTimer();
                //SetdelChangeTimer();
            }
            if (Ini.DriverCode == "LRK")
            {
                DLoadResults = new del(Lorak.LoadResults);
                DUploadOrders = new del(Lorak.UploadOrders);
            }
            if (Ini.DriverCode == "KDL")
            {
                DLoadDictionaries = new del(LabControl.KDL.LoadDitionaries);
               DLoadResults = new del(KDL.LoadResults);
                DUploadOrders = new del(KDL.UploadOrders);
            }
            if (Ini.DriverCode == "ARD")  //ЛИС Ариадна
            {
                DLoadDictionaries = new del(LabControl.Ariadna.LoadDictinaries);
                DLoadResults = new del(Ariadna.LoadResults);
                DUploadOrders = new del(Ariadna.UploadOrders);
            }

            if (Ini.DriverCode == "HLX")  //ЛИС Хеликс
            {
                DLoadResults = new del(Helix.LoadResults);
            }

            if (Ini.DriverCode == "HLF")  //ЛИС Хеликс
            {
                DLoadResults = new del(HelixF.LoadResults);
            }

            if (Ini.DriverCode == "TEST")  // 
            {
                DLoadResults = new del(Test.LoadResults);
            }

            if (Ini.DriverCode == "LITEX")  //
            {
                DLoadDictionaries = new del(Litex.LoadDictionaries);
            }

            SetOrderTimer(DUploadOrders);
            SetResultTimer(DLoadResults);
        }

        public static void NULL_Action()
        {
            Console.WriteLine("Неверно задан код драйвера!");
            Files.WriteLogFile("Неверно задан код драйвера!");
            Thread.Sleep(5000); Environment.Exit(1);
        }

        public static void NULL_DictAction()
        {
            Console.WriteLine("Справочники не будут загружены.");
        }

        //Ожидание ввода с консоли
        static public void WaitQ()
        {
            string message;
            StringComparer stringComparer = StringComparer.OrdinalIgnoreCase;

            Console.WriteLine("Введите Q для выхода из программы;");
            while (proceed)
            {
                message = Console.ReadLine();
                if (stringComparer.Equals("Q", message))
                {
                    proceed = false;
                }
                if (Ini.DriverCode == "SLS")
                    if (Ini.WriteBuffer) WriteToFileThread.Abort();
            }
        }
        
        //Таймер на получение результатов
        public static void SetResultTimer(del DLoadResults)
        {
            ResultTimer = new System.Timers.Timer(); //600000
            ResultTimer.Elapsed += (obj, e) => LoadResults(DLoadResults, e);
            if (Ini.GetTimer != 0)
            {
                ResultTimer.AutoReset = true;
                ResultTimer.Interval = Ini.GetTimer;
            }
            else
            {
                ResultTimer.AutoReset = false;
                ResultTimer.Interval = 1000;
            }
            ResultTimer.Enabled = false;
        }

        //Таймер на выгрузку заказов
        public static void SetOrderTimer(del DUploadOrders)
        {
            OrderTimer = new System.Timers.Timer(); //600000
            OrderTimer.Elapsed += (obj, e) => UploadOrders(DUploadOrders, e);
            if (Ini.UploadTimer != 0)
            {
                OrderTimer.AutoReset = true;
                OrderTimer.Interval = Ini.UploadTimer;
            }
            else
            {
                OrderTimer.AutoReset = false;
                OrderTimer.Interval = 2000;
            }
            OrderTimer.Enabled = false;
        }
        
        //Загрузка результатов
        static public void LoadResults(del LoadResults, ElapsedEventArgs e)
        {
            Console.WriteLine(e.SignalTime.ToString("HH:mm:ss")); 
            try
            {
                Console.WriteLine("Импорт результатов");
                LoadResults();
            }
            catch (Exception ex)
            {
                Files.WriteLogFile("Ошибка при импорте результатов: " +  ex.ToString()); Console.WriteLine(ex.Message);
            }
        }

        //Формирование и выгрузка заказов
        static public void UploadOrders(del DUploadOrders, ElapsedEventArgs e)
        {
            Console.WriteLine(e.SignalTime.ToString("HH:mm:ss"));
            try
            {
                Console.WriteLine("Экспорт заказов");
                DUploadOrders();
            }
            catch (Exception ex)
            {
                Files.WriteLogFile("Ошибка при экспорте заказов: " + ex.ToString()); Console.WriteLine(ex.Message);
                throw;
            }
        }

        //Таймер проверки изменений таблицы MSS_BARCODE_NUMBERS 
        public static void SetbcChangeTimer()
        {
            bcChangeTimer = new System.Timers.Timer(2000);
            
            bcChangeTimer.Elapsed += (obj, e) => Invitro.OrderCreate(obj, e);
            bcChangeTimer.Enabled = true;
        }

        //Таймер проверки изменений таблицы MSS_BARCODE_NUMBERS 
        public static void SetdelChangeTimer()
        {
            delChangeTimer = new System.Timers.Timer(3300);

            delChangeTimer.Elapsed += (obj, e) => Invitro.OrderDelete(obj, e);
            delChangeTimer.Enabled = true;
        }

        class ALab
        {
            //Загрузка результатов
            static public void LoadResults()
            {
                DataContext db = new DataContext(sqlConnectionStr);

                Table<Impdata> TImpdata = db.GetTable<Impdata>();
                Table<Restests> TRestests = db.GetTable<Restests>();
                Table<IdValues> TIdValue = db.GetTable<IdValues>();
                Table<Tests> TTests = db.GetTable<Tests>();
                Table<LAB_DEVS> TLAB_DEVS = db.GetTable<LAB_DEVS>();
                Table<OrderRequest> TOrder = db.GetTable<OrderRequest>();
                Table<Patient> TPatient = db.GetTable<Patient>();
                Table<PlExam> TPlExam = db.GetTable<PlExam>();
                Table<Patdirec> TPatdirec = db.GetTable<Patdirec>();
                Table<Motconsu> TMotconsu = db.GetTable<Motconsu>();
                Table<Dep> TDep = db.GetTable<Dep>();
                Table<MedDep> TMedDep = db.GetTable<MedDep>();
                
                //Получение ID для таблицы заказов Impdata
                var qIdValueI = (from idv in TIdValue
                                 where idv.KeyName == "Impdata"
                                 select idv).Single();
                int qId = qIdValueI.LastValue + 1;
                
                //Получение ID для таблицы заказов DS_RESTESTS
                var qIdValueR = (from idv in TIdValue
                                 where idv.KeyName == "DS_RESTESTS"
                                 select idv).Single();
                int qIdR = qIdValueR.LastValue + 1;
                
                //Определение пользователя и отделения 
                var qDep = (from de in TDep
                            join me in TMedDep on de.DepId equals me.DepId
                            where de.OrgId == Ini.OrgId
                            select new
                            {
                                DepId = me.DepId,
                                MedDepId = me.MedDepId,
                                MedecinsId = me.MedecinsId
                            }).SingleOrDefault();

                //Цикл по обновленным заказам в ЛИС
                bool existnew = true;
                bool done = true;
                bool removed = false;
                string updateversion;
                string Nr;
                string error;
                int? patid;
                DateTime resdate;
                int rowsAffected;

                while (existnew)
                {
                    ReqNextResults(out XDocument xdoc);

                    //Проверка сообщений об ошибках
                    if (xdoc.Element("Message").Attribute("Error") != null)
                    {   //Если есть ошибка, заканчиваем цикл
                        error = xdoc.Element("Message").Attribute("Error").Value;
                        Console.WriteLine(error);
                        existnew = false;
                    }
                    else
                    {   //Нет ошибки, продолжаем обработку
                        updateversion = xdoc.Element("Message").Element("Version").Attribute("Version").Value;
                        Nr = xdoc.Element("Message").Element("Referral").Attribute("Nr").Value;
                        done = Convert.ToBoolean(xdoc.Element("Message").Element("Referral").Attribute("Done").Value);
                        if (xdoc.Element("Message").Element("Referral").Attribute("Removed") != null)
                        {
                            removed = Convert.ToBoolean(xdoc.Element("Message").Element("Referral").Attribute("Removed").Value);
                            if (removed)
                            {
                                Console.WriteLine("Удаленный заказ...");
                                ImportConfirm(updateversion, Nr, out XDocument response);
                                Console.WriteLine(response);
                            }
                        }

                        //Загрузка из ЛИС в базу данных если заказ готов 
                        if (done)
                        {   //Выборка для вставки в Impdata
                            Console.WriteLine("Загрузка результатов в базу данных по заказу № {0}, версия выгрузки {1}", Nr, updateversion);
                            if (int.TryParse(xdoc.Element("Message").Element("Patient").Attribute("MisId").Value, out int patientsid)) { }
                            var qPatient = (from pe in TPatient
                                            where pe.Id == patientsid
                                            select pe.Id).Take(1);
                            Console.WriteLine("Пациент ID = {0}", patientsid);
                            if (qPatient.Count() == 0) { patid = null; Console.WriteLine("Пациент с ID = {0} не найден", patientsid); }
                            else { patid = patientsid; }
                            Console.WriteLine("Вставка в IMPDATA....");
                            resdate = Convert.ToDateTime(xdoc.Element("Message").Element("Referral").Attribute("DoneDate").Value);
                            var insetImpdata = from xe in xdoc.Elements("Message").Elements("Referral")
                                               select new Impdata
                                               {
                                                   ImpdataId = qId,
                                                   KEYCODE = xe.Attribute("Nr").Value,
                                                   Nom = xe.Parent.Element("Patient").Attribute("LastName").Value,
                                                   Prenom = xe.Parent.Element("Patient").Attribute("FirstName").Value + " "
                                                        + xe.Parent.Element("Patient").Attribute("MiddleName").Value,
                                                   Date_Consultation = Convert.ToDateTime(xe.Attribute("DoneDate").Value),
                                                   Date_Naissance = Convert.ToDateTime(xe.Parent.Element("Patient").Attribute("BirthDate").Value),
                                                   PATIENTS_ID = patientsid,
                                                   Mesure = Mesure
                                               };
                            TImpdata.InsertAllOnSubmit(insetImpdata); //Вставка в Impdata
                            qIdValueI.LastValue = qId;
                            db.SubmitChanges();
                            Console.WriteLine("Вставка в IMPDATA сделана");
                            Console.WriteLine("Вставка в DS_RESTESTS....");
                            //Выборка для вставки в DS_RESTESTS
                            var insetRestests = from xe in xdoc.Elements("Message").Elements("Blanks").Elements("Item").Elements("Tests").Elements("Item")
                                                select new Restests
                                                {
                                                    RestestsId = qIdR++,
                                                    ImpdataId = qId,
                                                    VAL = xe.Attribute("Value").Value,
                                                    RES_DATE = Convert.ToDateTime(xe.Attribute("ValueDate").Value),
                                                    UNIT = xe.Attribute("UnitName").Value,
                                                    STATE = (xe.Attribute("NormsFlag").Value == "1" ? 0 :
                                                            xe.Attribute("NormsFlag").Value == "0" ? 0 :
                                                            2),
                                                    MOTCONSU_ID = null,
                                                    COMMENT = xe.Attribute("Comment").Value,
                                                    LAB_METHODS_ID = (from me in TTests
                                                                      where me.GroupID == Ini.AnalyzerID && me.Code == xe.Attribute("Code").Value
                                                                      select me.LabMethodId ?? -1).Single(),
                                                    RES_TYPE = "D",
                                                    ITEM = null,
                                                    Rescode = xe.Attribute("GroupCode").Value,
                                                    NORM_TEXT_REC = xe.Attribute("Norms").Value,
                                                    NORM_STATE = (xe.Attribute("NormsFlag").Value == "1" ? 0 :
                                                                  xe.Attribute("NormsFlag").Value == "2" ? 1 :
                                                                  xe.Attribute("NormsFlag").Value == "4" ? 1 :
                                                                  xe.Attribute("NormsFlag").Value == "6" ? 1 :
                                                                  xe.Attribute("NormsFlag").Value == "3" ? 2 :
                                                                  xe.Attribute("NormsFlag").Value == "7" ? 2 :
                                                                    3),

                                                };
                            TRestests.InsertAllOnSubmit(insetRestests); //Всатвка в DS_RESTESTS
                                                                        //Подсчет микроорганизмов в исследовании
                            int itemR = 1;
                            var itemCount = from xe in xdoc.Elements("Message").Elements("Blanks").Elements("Item").Elements("Tests").Elements("Item")
                                                .Elements("Microorganisms").Elements("Item")
                                            select new
                                            {
                                                itemN = itemR++,
                                                TestName = xe.Parent.Parent.Attribute("Name").Value,
                                                MicroorganismName = xe.Attribute("Name").Value,
                                                Quantity = xe.Attribute("Value").Value
                                            };
                            //Цикл по микрорганизмам
                            foreach (var item in itemCount)
                            {   //Название исследования 
                                Restests rTN = new Restests
                                {
                                    RestestsId = qIdR++,
                                    ImpdataId = qId,
                                    VAL = item.TestName,
                                    RES_DATE = DateTime.Now,
                                    LAB_METHODS_ID = (from me in TTests
                                                      join ge in TLAB_DEVS on me.GroupID equals ge.LabDevsId
                                                      where ge.Label == "Drugs" && me.Code == "Name"
                                                      select me.LabMethodId ?? -1).Single(),
                                    ITEM = item.itemN,
                                    STATE = 0,
                                    MOTCONSU_ID = null,
                                };
                                TRestests.InsertOnSubmit(rTN);
                                //Название микроорганизма
                                Restests rMN = new Restests
                                {
                                    RestestsId = qIdR++,
                                    ImpdataId = qId,
                                    VAL = item.MicroorganismName,
                                    RES_DATE = DateTime.Now,
                                    LAB_METHODS_ID = (from me in TTests
                                                      join ge in TLAB_DEVS on me.GroupID equals ge.LabDevsId
                                                      where ge.Label == "Drugs" && me.Code == "Microorganism"
                                                      select me.LabMethodId ?? -1).Single(),
                                    ITEM = item.itemN,
                                    STATE = 0,
                                    MOTCONSU_ID = null,
                                };
                                TRestests.InsertOnSubmit(rMN);
                                //Количество
                                Restests rQ = new Restests
                                {
                                    RestestsId = qIdR++,
                                    ImpdataId = qId,
                                    VAL = item.Quantity,
                                    RES_DATE = DateTime.Now,
                                    LAB_METHODS_ID = (from me in TTests
                                                      join ge in TLAB_DEVS on me.GroupID equals ge.LabDevsId
                                                      where ge.Label == "Drugs" && me.Code == "Quantity"
                                                      select me.LabMethodId ?? -1).Single(),
                                    ITEM = item.itemN,
                                    STATE = 0,
                                    MOTCONSU_ID = null,
                                };
                                TRestests.InsertOnSubmit(rQ);
                                //Выборка антибиотиков для вставки
                                insetRestests = from xe in xdoc.Elements("Message").Elements("Blanks").Elements("Item").Elements("Tests").Elements("Item")
                                                    .Elements("Microorganisms").Elements("Item").Elements("Drugs").Elements("Item")
                                                select new Restests
                                                {
                                                    RestestsId = qIdR++,
                                                    ImpdataId = qId,
                                                    VAL = xe.Attribute("Value").Value,
                                                    RES_DATE = Convert.ToDateTime(xe.Parent.Parent.Parent.Parent.Parent.Parent.Attribute("DoneDate").Value),
                                                    STATE = 0,
                                                    MOTCONSU_ID = null,
                                                    LAB_METHODS_ID = (from me in TTests
                                                                      join ge in TLAB_DEVS on me.GroupID equals ge.LabDevsId
                                                                      where ge.Label == "Drugs" && me.Code == xe.Attribute("Code").Value
                                                                      select me.LabMethodId ?? -1).Single(),
                                                    RES_TYPE = "D",
                                                    ITEM = item.itemN,
                                                };
                                TRestests.InsertAllOnSubmit(insetRestests);
                            }
                            qIdValueR.LastValue = qIdR - 1;
                            db.SubmitChanges();
                            Console.WriteLine("Вставка в DS_RESTESTS сделана");

                            //Редактирование справочников лабораторного модуля для представления результатов
                            //Создание видов исследований, если нет в справочнике
                            var qDictGgoups = (from xe in xdoc.Elements("Message").Elements("Blanks").Elements("Item").Elements("Tests").Elements("Item")
                                               select new
                                               {
                                                   GName = xe.Attribute("GroupName").Value,
                                                   GCode = xe.Attribute("GroupCode").Value
                                               }).Distinct();
                            foreach (var elm in qDictGgoups)
                            {
                                Console.WriteLine("{0}  {1}", elm.GCode, elm.GName);
                                rowsAffected = db.ExecuteCommand(@"declare @id_dct int, @id_rel int
                            if not exists( select Num from DS_OUTER_DICTPARAMS where NUM = {0})
                            begin
	                            exec dbo.[up_get_id] 'DS_OUTER_DICTPARAMS',1,@id_dct output
	                            insert DS_OUTER_DICTPARAMS (DS_OUTER_DICTPARAMS_ID,LABEL, NUM, DS_OUTER_DICT_ID, IS_TOPLEVEL,SHOWINRESUME, HIDE_RESCOUNTRPT, USECODE )
	                            values (@id_dct, {1}, {0}, 3, 1, 1, 0, 0)

	                            exec dbo.[up_get_id] 'DS_OUTER_RELATIONS',1,@id_rel output
	                            insert into DS_OUTER_RELATIONS  (DS_OUTER_RELATIONS_ID,PARENT_ID,CHILD_ID,REL_LEVEL, IS_PRIMARY)
	                            values(@id_rel, {2}, @id_dct,0,1)
                            end", elm.GCode, elm.GName, Ini.ResGroupId);
                            }
                            //Создание параметров, если их нет в справочнике
                            rowsAffected = db.ExecuteCommand(@"declare	@count int, @id_par int, @id_bio int, @id_out int, @id_dct int, @id_rel int, @i int = 1	
                                declare @Labels table(
				                                NUM int,
				                                LAB_METHODS_ID int,
				                                LABEL varchar(250),
				                                CODE varchar(50),
				                                RESCODE varchar(50)
				                                )
                                insert @Labels
                                select ROW_NUMBER() over(order by dr.LAB_METHODS_ID), dr.LAB_METHODS_ID,  lm.LABEL, lm.CODE, dr.RESCODE 
                                from DS_RESTESTS dr
                                join LAB_METHODS lm on dr.LAB_METHODS_ID=lm.LAB_METHODS_ID
                                left outer join LAB_METHODBIO lmb on dr.LAB_METHODS_ID=lmb.LAB_METHODS_ID
                                where dr.IMPDATA_ID = {0} and lmb.LAB_METHODBIO_ID is null

                                set @count = (select count(LAB_METHODS_ID) from @Labels)

	                            exec dbo.[up_get_id] 'DS_OUTER_PARAMS',@count,@id_out output
	                            insert DS_OUTER_PARAMS (DS_OUTER_PARAMS_ID,LABEL,CODE)
	                            select NUM - 1 + @id_out, LABEL,CODE 
	                            from @Labels
		
	                            exec dbo.[up_get_id] 'DS_PARAMS',@count,@id_par output
	                            insert DS_PARAMS (DS_PARAMS_ID,LABEL, DS_OUTER_PARAMS_ID)
	                            select NUM - 1 + @id_par, l.LABEL, p.DS_OUTER_PARAMS_ID 
	                            from @Labels l
	                            join DS_OUTER_PARAMS p on l.CODE=p.CODE 
	                            where p.DS_OUTER_PARAMS_ID>=@id_out

	                            exec dbo.[up_get_id] 'LAB_METHODBIO',@count,@id_bio output
	                            insert into LAB_METHODBIO (LAB_METHODBIO_ID,LAB_METHODS_ID,DS_PARAMS_ID)
	                            select NUM - 1 + @id_bio, LAB_METHODS_ID, p.DS_PARAMS_ID  
	                            from @Labels l
	                            join DS_OUTER_PARAMS pout on l.CODE=pout.CODE
	                            join DS_PARAMS p on pout.DS_OUTER_PARAMS_ID=p.DS_OUTER_PARAMS_ID 
	                            where pout.DS_OUTER_PARAMS_ID>=@id_out

	                            exec dbo.[up_get_id] 'DS_OUTER_DICTPARAMS',@count,@id_dct output
	                            insert into DS_OUTER_DICTPARAMS (DS_OUTER_DICTPARAMS_ID,LABEL,DS_OUTER_DICT_ID,DS_OUTER_PARAMS_ID,IS_TOPLEVEL,SHOWINRESUME)
	                            select NUM - 1 + @id_dct, l.LABEL, 3, p.DS_OUTER_PARAMS_ID, 0, 0 
	                            from @Labels l
	                            join DS_OUTER_PARAMS p on l.CODE=p.CODE 
	                            where p.DS_OUTER_PARAMS_ID>=@id_out
	
	                            exec dbo.[up_get_id] 'DS_OUTER_RELATIONS',@count,@id_rel output
                                while (@i<=@count)
                                Begin	
	                                insert into DS_OUTER_RELATIONS  (DS_OUTER_RELATIONS_ID,PARENT_ID,CHILD_ID,REL_LEVEL, IS_PRIMARY)
	                                select l.NUM - 1 + @id_rel, dic.DS_OUTER_DICTPARAMS_ID, dp.DS_OUTER_DICTPARAMS_ID, 0,1
	                                from @Labels l
	                                join DS_OUTER_PARAMS p on l.CODE=p.CODE
	                                join DS_OUTER_DICTPARAMS dp on p.DS_OUTER_PARAMS_ID = dp.DS_OUTER_PARAMS_ID
	                                join DS_OUTER_DICTPARAMS dic on l.RESCODE = dic.NUM
	                                where l.NUM = @i and p.DS_OUTER_PARAMS_ID>=@id_out 
	                                set @i=@i+1
                                continue
                                end
                         ", qId);

                            //Подтверждение импорта 
                            ImportConfirm(updateversion, Nr, out XDocument response);
                            
                            //Обновление статуса заказа 
                            rowsAffected = db.ExecuteCommand(
                            @"update MSS_ORDER_REQUEST set STATE=3
                        where MSS_ORDER_REQUEST_ID = {0}", qId);
                            //Удаление ранее импортированных результатов 
                            rowsAffected = db.ExecuteCommand(@"declare @count int
	                    set @count =(SELECT  COUNT(*) 
		                FROM IMPDATA impdata
		                join IMPDATA imp on imp.KEYCODE=impdata.KEYCODE and imp.PATIENTS_ID=impdata.PATIENTS_ID 
				            and imp.Date_Naissance >= IMPDATA.Date_Naissance 
				            and isnull(imp.Mesure,'')=isnull(impdata.Mesure,'') 
		                where impdata.Impdata_ID={0})

	                    if @count >1 
	                    begin
		                    update IMPDATA set ImpDeleted=1 
		                    where Impdata_ID in (select top (@count-1) imp.Impdata_ID
			                    FROM IMPDATA impdata
			                    join IMPDATA imp on imp.KEYCODE=impdata.KEYCODE and imp.PATIENTS_ID=impdata.PATIENTS_ID 
				                    and imp.Date_Naissance >= IMPDATA.Date_Naissance 
				                    and isnull(imp.Mesure,'')=isnull(impdata.Mesure,'') 
			                    where impdata.Impdata_ID={0}
		                    order by imp.Impdata_ID )

		                    update DS_RESTESTS set STATE=-1 
		                    where Impdata_ID in ( select top (@count-1) imp.Impdata_ID
			                    FROM IMPDATA impdata
			                    join IMPDATA imp on imp.KEYCODE=impdata.KEYCODE and imp.PATIENTS_ID=impdata.PATIENTS_ID 
				                    and imp.Date_Naissance >= IMPDATA.Date_Naissance 
				                    and isnull(imp.Mesure,'')=isnull(impdata.Mesure,'') 
			                    where impdata.Impdata_ID={0}
		                    order by imp.Impdata_ID )
	                    end", qId);
                            //Отправка в историю болезни (создание записей в таблице MOTCONSU)
                            //Определение события
                            int? qEventId = (from oe in TOrder
                                             join pe in TPatdirec on oe.OrderId equals pe.Id
                                             join me in TMotconsu on pe.MotconsuId equals me.Id
                                             where oe.OrderNumber == Convert.ToInt32(Nr)
                                             select me.EventId).SingleOrDefault();
                            if (qEventId == null) qEventId = 0;
                            Console.WriteLine("Создание записи в истории заболевания");
                            rowsAffected = db.ExecuteCommand(@"
                           	declare @m_id int, @DATEtime datetime = Getdate();
		                    set	@m_id=(select top 1 m.MOTCONSU_ID from MOTCONSU m
		                    where PATIENTS_ID={1} 
		                    and cast(m.DATE_CONSULTATION as DATE)=cast({2} as date)
		                    and m.MODELS_ID={3} and m.MEDECINS_ID={4})
		
		                   if @m_id is null 
		                     begin
			                    exec up_get_id 'MOTCONSU',1,@m_id out
			                    INSERT into MOTCONSU (MOTCONSU_ID,    PATIENTS_ID,  MEDECINS_ID, FM_DEP_ID,MEDECINS_CREATE_ID, 
			                    MODELS_ID, DATE_CONSULTATION,  CREATE_DATE_TIME, REC_STATUS, MODIFY_DATE_TIME, MEDECINS_MODIFY_ID, KRN_MODIFY_USER_ID, 
			                    MOTCONSU_EV_ID, REC_NAME, MEDDEP_ID,
			                    EV_GOSP, EV_CLOSE,PUBLISHED,CHANGED, POVTORN_J_PRIEM,
			                    RASHIRENN_J_ST_PRAESENS,BOL_NIHN_J_LIST,  DISPANSERN_J_UHET, 
			                    SANATORNO_KURORTNAQ_SPRAV,  
			                    ENDOSKOPIQ_METOD, BIOPSIQ_FEGDS, BIOPSIQ_ZNDOSKOPIQ, BIOPSIQ_GINEKOLOG, BIOPSIQ_IRURG,
			                    NEXT_PATIENTS,TIME_PATIENTS) 
			                    values (@m_id, {1} , {4}, {5}, {4}, {3}, {2}, @DATEtime,'W', @DATEtime, {4}, {4}, 
			                    case when {7}=0 then null else {7} end, 'Лаборатория', {6}, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
					                    1, @DATEtime)
		                    end
		                 --Обновление IMPDATA
		                 update IMPDATA set STATE=1
		                 where Impdata_ID={0} 
		                 --Обновление DS_RESTESTS
		                 update DS_RESTESTS set MOTCONSU_ID = @m_id
		                 where Impdata_ID={0}"
                             , qId, patid, resdate, Ini.ModelId, qDep.MedecinsId, qDep.DepId, qDep.MedDepId, qEventId.Value);
                            qId++; //Impdata_ID
                        }
                    } //Нет ошибки в передаче результатов
                      //Console.WriteLine("Ждем две минуты"); Thread.Sleep(120000);
                    existnew = false;
                }
                existnew = false;
                qIdValueI.LastValue = qId - 1;
                qIdValueR.LastValue = qIdR - 1;

                db.SubmitChanges();
                db.Dispose();
            }
            //Формирование и выгрузка заказов
            static public void UploadOrders()
            {
                DataContext db = new DataContext(sqlConnectionStr); //Создание контекста для работы с базой данных
                                                                    //Создание объектов для контекста
                Table<Patient> TPatient = db.GetTable<Patient>();
                Table<Patdirec> TPatdirec = db.GetTable<Patdirec>();
                Table<DirServ> TDirServ = db.GetTable<DirServ>();
                Table<OrderRequest> TOrder = db.GetTable<OrderRequest>();
                Table<BarcodeNum> TBarcodeNum = db.GetTable<BarcodeNum>();
                Table<ServofOrder> TServofOrder = db.GetTable<ServofOrder>();
                Table<IdValues> TIdValue = db.GetTable<IdValues>();
                Table<Profiles> TProfiles = db.GetTable<Profiles>();
                Table<Biomaterial> TBiomaterial = db.GetTable<Biomaterial>();
                Table<ProfilesBiomaterial> TProfileBiomaterial = db.GetTable<ProfilesBiomaterial>();
                Table<ProfileServ> TProfileServ = db.GetTable<ProfileServ>();
                //Выбор пациентов с направлениями во внешнюю лабораторию 
                Console.WriteLine("Создание списка пациентов для отправки заказов");
                var qReqOrder = (from be in TBarcodeNum
                                 join pe in TPatdirec on be.PatdirId equals pe.Id
                                 where DateTime.Now.AddHours(-24) < Convert.ToDateTime(pe.DateBio)
                                          && DateTime.Now.AddMinutes(-5) > Convert.ToDateTime(pe.DateBio)
                                 orderby be.BarcodeId
                                 select pe.PatientID).Distinct()
                     .Except(
                     from re in TOrder
                     join pe in TPatdirec on re.OrderId equals pe.OrderId
                     join be in TBarcodeNum on pe.Id equals be.PatdirId
                     where DateTime.Now.AddHours(-24) < Convert.ToDateTime(pe.DateBio)
                              && DateTime.Now.AddMinutes(-5) > Convert.ToDateTime(pe.DateBio)
                     select pe.PatientID
                     );
                //Получение ID для табицы заказов MSS_ORDER_REQUEST
                var qIdValue = (from idv in TIdValue
                                where idv.KeyName == "MSS_ORDER_REQUEST"
                                select idv).Single();
                int qId = qIdValue.LastValue;

                int rowsAffected;
                // Цикл по PTIENTS_ID для создания заказов 
                foreach (var item in qReqOrder)
                {
                    qId++; //Следующий ID для MSS_ORDER_REQUEST
                    qIdValue.LastValue = qId;  //Обновление ID таблицы заказов
                    db.SubmitChanges();   //Подтверждение изменений в базе данных
                                          //Запуск скрипта для создания строки заказа и обновления направлений во внешнюю лабораторию
                    Console.WriteLine("Создание заказа для пациента ID = {0}", item);
                    rowsAffected = db.ExecuteCommand(
                        @" declare @order_counter varchar(20), @ord_num int
	
	                set @order_counter='order_num'+'_'+  cast({1} as varchar(5))
	                exec mss_counter @order_counter, @ord_num out
                    declare @ext_code varchar(20) = (select top 1 CODE_AN from FM_ORG where FM_ORG_ID={1})

	                insert MSS_ORDER_REQUEST (MSS_ORDER_REQUEST_ID, ORDER_NUMBER, FM_ORG_ID, PATIENTS_ID, DATE_ORDER, STATE)
	                values ({2}, @ext_code + substring(cast( @ord_num as varchar(20)), 2,20), 
                    {1}, {0}, GETDATE(), 0)

	                update PATDIREC set MSS_ORDER_REQUEST_ID={2}
	                from PATDIREC pd
	                join DIR_SERV dsr on dsr.PATDIREC_ID=pd.PATDIREC_ID
	                join MSS_ORDER_SERV osr on osr.FM_SERV_ID=dsr.FM_SERV_ID
	                join MSS_BARCODE_NUMBERS brc on brc.PATDIREC_ID=pd.PATDIREC_ID
	                where pd.PATIENTS_ID={0} and datediff(HH,pd.DATE_BIO, GETDATE())<=24 
		                and  pd.MSS_ORDER_REQUEST_ID is null
		                and isnull(osr.state,0)=1 "
                    , item, Ini.OrgId, qId);
                    //Данные о пациенте
                    var qPatient = (from re in TOrder
                                    join pt in TPatient on re.PatientId equals pt.Id
                                    where re.OrderId == qId
                                    select new
                                    {
                                        MisId = pt.Id,
                                        LastName = pt.LastName,
                                        FirstName = pt.FirstName,
                                        MiddleName = pt.MiddleName,
                                        Gender = (pt.Pol == 0 ? 1 : pt.Pol == 1 ? 2 : 3),
                                        BirthDate = pt.BirthDate
                                    }).Single();
                    //Данные о заказе
                    var qReferral = (from re in TOrder
                                     where re.OrderId == qId
                                     select new
                                     {
                                         MisId = re.OrderId,
                                         Nr = re.OrderNumber,
                                         Date = re.DateOrder,
                                         SamplingDate = re.DateOrder
                                     }).Single();
                    //Данные о пробирках 
                    var qAssays = (from re in TOrder
                                   join pd in TPatdirec on re.OrderId equals pd.OrderId
                                   join bc in TBarcodeNum on pd.Id equals bc.PatdirId
                                   join bm in TBiomaterial on bc.BiotypeId equals bm.LabBiotypeId
                                   where re.OrderId == qId
                                   select new
                                   {
                                       Barcode = bc.BarcodeNumber,
                                       BiomaterialCode = bm.Code,
                                       BarcodeId = bc.BarcodeId
                                   }).Distinct();
                    //Данные о исследованиях
                    var qOrder = (from re in TOrder
                                  join pd in TPatdirec on re.OrderId equals pd.OrderId
                                  join bc in TBarcodeNum on pd.Id equals bc.PatdirId
                                  join bm in TBiomaterial on bc.BiotypeId equals bm.LabBiotypeId
                                  join ds in TDirServ on pd.Id equals ds.PatdirId
                                  join ps in TProfileServ on ds.ServId equals ps.ServId
                                  join pe in TProfiles on ps.ProfileId equals pe.ProfileId
                                  where re.OrderId == qId
                                  select new
                                  {
                                      Code = pe.Code,
                                      BarcodeId = bc.BarcodeId
                                  });

                    CultureInfo culture = new CultureInfo("ru-RU"); //формат преобразования даты
                                                                    //Создание тела сообщения для создания заказа в ЛИС 
                    XElement messagebody = new XElement("Root"
                            , new XElement("Patient",
                                           new XAttribute("MisId", qPatient.MisId),
                                           new XAttribute("LastName", qPatient.LastName),
                                           new XAttribute("FirstName", qPatient.FirstName),
                                           new XAttribute("MiddleName", qPatient.MiddleName),
                                           new XAttribute("Gender", qPatient.Gender),
                                           new XAttribute("BirthDate", Convert.ToString(qPatient.BirthDate, culture))
                                          )
                            , new XElement("Referral",
                                          new XAttribute("MisId", qReferral.MisId),
                                          new XAttribute("Nr", qReferral.Nr),
                                          new XAttribute("Date", Convert.ToString(qReferral.Date, culture)),
                                          new XAttribute("SamplingDate", Convert.ToString(qReferral.SamplingDate, culture))
                                          )
                            , new XElement("Assays",
                                from res in qAssays
                                select new XElement("Item",
                                                    new XAttribute("Barcode", res.Barcode),
                                                    new XAttribute("BiomaterialCode", res.BiomaterialCode),
                                                    new XElement("Orders",
                                                                 from ord in qOrder
                                                                 where ord.BarcodeId == res.BarcodeId
                                                                 select new XElement("Items",
                                                                                     new XAttribute("Code", ord.Code)
                                                                                    )
                                                                )
                                                   )
                                          ) //Assay
                            );
                    //Обновление статуса заказа 
                    rowsAffected = db.ExecuteCommand(
                        @"update MSS_ORDER_REQUEST set STATE=1
                    where MSS_ORDER_REQUEST_ID = {0}"
                        , qId);
                    //Формирование сообщения для заказа и отправка заказа
                    Message.ALab result = new Message.ALab();
                    if (result.Request("query-create-referral", messagebody, out XDocument responce, out string errmessage)) Console.WriteLine("Создание заказа в ЛИС...");
                    else Console.WriteLine("Ошибка при формировании или выгрузке заказа");
                    string LISNumber = responce.Element("Message").Element("Referral").Attribute("LisId").Value;
                    string MISNumber = responce.Element("Message").Element("Referral").Attribute("MisId").Value;
                    if (LISNumber != "")
                    {   //Обновление статуса заказа 
                        rowsAffected = db.ExecuteCommand(
                        @"update MSS_ORDER_REQUEST set STATE=2
                    where MSS_ORDER_REQUEST_ID = {0}"
                        , qId);
                        Console.WriteLine("Заказ успешно отправлен в ЛИС! Номер заказа в Медиалоге {0}, Номер заказа в ЛИС {1}", MISNumber, LISNumber);
                    }
                    db.SubmitChanges();   //Подтверждение изменений в базе данных
                }
                db.Dispose();
                //Очистка контекста работы с базой данных
            }
            //Удаление заказа из ЛИС
            static public void DeleteOrder(string MisNum)
            {
                XElement messagebody = new XElement("Root"
                    , new XElement("Query"
                                  , new XAttribute("MisId", "")
                                  , new XAttribute("Nr", MisNum)
                                  , new XAttribute("LisId", "")
                                 )
                     );

                Message.ALab delete = new Message.ALab();
                if (delete.Request("query-referral-remove", messagebody, out XDocument responce, out string errmessage)) Console.WriteLine(responce);

            }
            //Запрос на получение результатов по заказу 
            static public void ReqResults(string MisNum, out XDocument resultresponce)
            {
                XElement messagebody = new XElement("Root"
                    , new XElement("Query"
                                  , new XAttribute("MisId", "")
                                  , new XAttribute("Nr", MisNum)
                                  , new XAttribute("LisId", "")
                                 )
                     );
                Message.ALab result = new Message.ALab();
                if (result.Request("query-referral-results", messagebody, out resultresponce, out string errmessage)) Console.WriteLine("Выгрузка результата из ЛИС");
            }
            //Запрос на получение результатов следующему направлению
            static public void ReqNextResults(out XDocument resultresponce)
            {
                XElement messagebody = new XElement("Root");

                Message.ALab result = new Message.ALab();
                if (result.Request("query-next-referral-results", messagebody, out resultresponce, out string errmessage)) Console.WriteLine("Выгрузка результата из ЛИС");
            }
            //Запрос на подтверждение импорта по следующему направлению
            static public void ImportConfirm(string updateversion, string MisNum, out XDocument resultresponce)
            {
                 XElement messagebody =
                  new XElement("Root"
                      , new XElement("Version", new XAttribute("Version", updateversion))
                      , new XElement("Referral"
                        , new XAttribute("MisId", "")
                        , new XAttribute("Nr", MisNum)
                        , new XAttribute("LisId", "")
                                    )
                               );
                Message.ALab result = new Message.ALab();
                if (result.Request("result-referral-results-import", messagebody, out resultresponce, out string errmessage)) Console.WriteLine("Импорт по заказу №{0} подтвержден", MisNum);
            }
        }

        class CLD
        {
            //Загрузка результатов в формате XML
            static public void LoadResults()
            {
                XDocument xdoc; 
                int p_id = 0;

                string[] resfiles = Directory.GetFiles(Ini.ResultFilePath, "*.xml");
                foreach (string filepath in resfiles)
                {
                    DataContext db = new DataContext(sqlConnectionStr);
                    Table<Impdata> TImpdata = db.GetTable<Impdata>();
                    
                    xdoc = XDocument.Load(filepath);
                    Console.WriteLine("Загрузка результатов из файла {0}", filepath);
                    string StrPatID = xdoc.Element("Request").Element("AmbulantCard").Value;

                    bool ok = true;
                    bool attachPdf = Directory.Exists(Ini.PdfFilePath) && (Ini.RubricsId != 0);

                    if (!DBControl.GetPatientID(db, StrPatID, filepath, out p_id))
                    {
                        ok = false;
                        PatientName Name = new PatientName
                        {
                            LastName = xdoc.Element("Request").Element("Patient").Element("LastName").Value,
                            FirstName = xdoc.Element("Request").Element("Patient").Element("FirstName").Value,
                            MiddleName = xdoc.Element("Request").Element("Patient").Element("MiddleName").Value
                        };
 
                        int BirthYear = 1990;
                        try { BirthYear = Convert.ToInt32(xdoc.Element("Request").Element("Patient").Element("BirthYear").Value); }
                        catch { }

                        PatientBirthDate BirthDate = new PatientBirthDate
                        {
                            BirthYear = BirthYear,
                            BirthDay = 0,
                            BirthMonth = 0
                        };

                        ok = DBControl.GetPatientID(db, Name, BirthDate, filepath, out p_id);
                    }

                    string KeyCode = xdoc.Element("Request").Element("RequestCode") != null ? 
                                    xdoc.Element("Request").Element("RequestCode").Value
                                  : xdoc.Element("Request").Element("OrderNumber").Value;    //OrderNumber
                    Console.WriteLine("Создание записи в Impdata");
                    bool NewImp = DBControl.GetImpdataID(db, KeyCode, out int ImpdataId);
                    Console.WriteLine("Запись создана");
                    //Подсчет количесва параметров исследований
                    int  resCount = (from xe in xdoc.Elements("Request").Elements("SampleResults").Elements("SampleResult").
                                        Elements("TargetResults").Elements("TargetResult").Elements("Works").Elements("Work")
                                         select xe.Element("Code")).Count();
                    
                    DateTime resdate = Convert.ToDateTime(xdoc.Element("Request").Element("ResultDate").Value);

                    if (NewImp)
                    {
                        Console.WriteLine("Количетсво строк {0}, id записи {1}", resCount, ImpdataId);
                        var insetImpdata = from xe in xdoc.Elements("Request")
                                           select new Impdata
                                           {
                                               ImpdataId = ImpdataId,
                                               KEYCODE = KeyCode,
                                               Nom = xe.Element("Patient").Element("LastName").Value,
                                               Prenom = xe.Element("Patient").Element("FirstName").Value + " " + xe.Element("Patient").Element("MiddleName").Value,
                                               Date_Naissance = new DateTime(Convert.ToInt32(xe.Element("Patient").Element("BirthYear").Value), 1, 1),
                                               Date_Consultation = resdate,
                                               //PATIENTS_ID = p_id == 0 ? null : p_id,
                                               Mesure = Mesure
                                           };
                        //Вставка в Impdata
                        TImpdata.InsertAllOnSubmit(insetImpdata);
                        db.SubmitChanges();
                    }

                    List<RestestData> results = new List<RestestData>();
                    
                    //Комментарии к исследованию
                    XElement commentsX = xdoc.Element("Request").Element("SampleResults").Element("SampleResult").
                                                     Element("TargetResults").Element("TargetResult").Element("Comments");

                    if (commentsX != null)
                    {
                        int methodId = DBControl.LabMethodCheck(db, "AnaComment");
                        //Создание справочников
                        if (methodId == 0)
                        {
                            if (Ini.ResGroupId == 0)
                                DBControl.BuildMethods(db, "Комментарии к исследованию", "AnaComment", "", "", out methodId);
                            else
                                DBControl.BuildDitionary(db, "Комментарии к исследованию", "AnaComment", "", "", Ini.ResGroupId, out methodId);
                        }
                        string val = commentsX.Value;
                        if (val != string.Empty)
                        {
                            //Компановка результатов
                            results.Add(new RestestData
                            {
                                VAL = val,
                                MethodId = methodId
                            }
                                    );
                        }
                    }

                    //Выборка для вставки в DS_RESTESTS
                    var qRestests = xdoc.Elements("Request").Elements("SampleResults").Elements("SampleResult").
                                                     Elements("TargetResults").Elements("TargetResult").Elements("Works").Elements("Work").
                                    Select( xe => new
                                    {
                                        VAL = xe.Element("Value").Value,
                                        UNIT = xe.Element("UnitName").Value,
                                        Rescode = xe.Element("Code").Value,
                                        ResName = xe.Element("Name").Value,
                                        NORM_TEXT_REC = xe.Element("Norm").Element("Norms").Value
                                    });

 
                    foreach (var q in qRestests)
                    {
                        int methodId = DBControl.LabMethodCheck(db, q.Rescode);
                        //Создание справочников
                        if (methodId == 0)
                        {
                            if (Ini.ResGroupId == 0)
                                DBControl.BuildMethods(db, q.ResName, q.Rescode, q.Rescode, q.UNIT, out methodId);
                            else
                               DBControl.BuildDitionary(db, q.ResName, q.Rescode, q.Rescode, q.UNIT, Ini.ResGroupId, out methodId);
                        }
                        results.Add(new RestestData
                        {
                            VAL = q.VAL,
                            UNIT = q.UNIT,
                            NORM_TEXT_REC = q.NORM_TEXT_REC,
                            MethodId = methodId
                        });
                    }

                    //Вставки в DS_RESTESTS
                    DBControl.InsertRestestDataIDM(db, results, ImpdataId, resdate);

                    //Отправка результатов на тестового пациента
                    if (Ini.TestPatientID != 0) { ok = true; p_id = Ini.TestPatientID; }

                    if (ok)
                    {
                        DBControl.UpdateImpdata(db, ImpdataId, p_id);
                        int motconsuId = DBControl.InsertToEMCe(ImpdataId, p_id, resdate, NewImp);
                        DBControl.InsetToUserTables(ImpdataId);
                        DBControl.UpdateDirAnsw(ImpdataId);

                        string pdfFilePath = Ini.ResultFilePath.Replace("xml", "pdf");
                        string fileNamePart = Path.GetFileName(filepath).Replace(".xml", "");
                        string[] pdffilepathes = Directory.GetFiles(pdfFilePath, fileNamePart + "*.pdf");

                        //Обработка pdf файла и прикрепление к ЭМК
                        if (attachPdf && motconsuId != 0)
                        {
                            string folder = string.Format("{0}-{1}/{2}", resdate.Year, Convert.ToString(resdate.Month).PadLeft(2, '0'), p_id);
                            string pdfDirectory = string.Format("{0}/{1}", Ini.PdfFilePath, folder);

                            if (!Directory.Exists(pdfDirectory))
                            {
                                DirectoryInfo dir = Directory.CreateDirectory(pdfDirectory);
                            }

                            foreach (string pdffp in pdffilepathes)
                            {
                                string medialogfp = string.Format("{0}/{1}", pdfDirectory, Path.GetFileName(pdffp));
                                if (File.Exists(medialogfp))
                                {
                                    File.Delete(medialogfp);
                                    DBControl.DeletePdfRefrense(p_id, Path.GetFileName(pdffp));
                                }
                                if (File.Exists(pdffp))
                                {
                                    File.Copy(pdffp, medialogfp);
                                    File.Delete(pdffp);
                                    DBControl.AttachPdf(motconsuId, p_id, Path.GetFileName(pdffp), resdate, folder, Processing.Ini.RubricsId);
                                }
                            }
                        }

                    }
                    db.Dispose();

                    if (Directory.Exists(Ini.ArchiveFilePath))
                    {
                        string fp = string.Format("{0}/{1}", Ini.ArchiveFilePath, Path.GetFileName(filepath));
                        if (File.Exists(fp)) File.Delete(fp);
                        File.Copy(filepath, fp);
                        File.Delete(filepath);
                    }
                    else
                    if (Ini.DeleteDownloaded)
                        File.Delete(filepath);
                }
            }

            //Формирование и выгрузка заказов в формате XML
            static public void UploadOrders()
            {
                CultureInfo culture = new CultureInfo("ru-RU"); //формат преобразования даты
                List<Request> RequestInfoList = new List<Request>();

                //DBControl.CreateOrders(out RequestInfoList);
                DBControl.GetOrdersList(0, 5, out RequestInfoList);

                Console.WriteLine("Колво {0}", RequestInfoList.Count);
                PatientInfo qPatient = new PatientInfo();
                List<AssayInfo> Assay = new List<AssayInfo>();
                List<Order> qOrder = new List<Order>();
                if ( RequestInfoList.Count == 0) Console.WriteLine("Нет направлений для заказов");
                else
                {
                    foreach (Request qRequest in RequestInfoList)
                    {
                        //Получение данных заказа из базы данных
                        switch (Ini.OrderOption)
                        {
                            //таблица DATA286 (Гинекологический статус) 
                            case 1: DBControl.GetOrderInfoV4_G2(qRequest.ID, out qPatient, out Assay, out qOrder);
                                break;
                            //без использования таблицы Гинекологический статус
                            case 2: DBControl.GetOrderInfoV4_1(qRequest.ID, out  qPatient, out  Assay, out  qOrder);
                                break;
                            //старый вариант без множественной привязки услуг к биоматериалам и без использования таблицы Гинекологический статус 
                            case 3:
                                DBControl.GetOrderInfoV3(qRequest.ID, out qPatient, out Assay, out qOrder);
                                break;
                            //таблица DATA_GYNAECOL_STATUS (Гинекологический статус) 
                            case 0: DBControl.GetOrderInfoV4(qRequest.ID, out qPatient, out Assay, out qOrder);
                                break;
                        }
                        //DBControl.GetOrderInfoV4_G2(qRequest.ID, out PatientInfo qPatient, out List<AssayInfo> Assay, out List<Order> qOrder);
                        //DBControl.GetOrderInfoV4(qRequest.ID, out  qPatient, out  Assay, out qOrder);
                        //DBControl.GetOrderInfoV3(qRequest.ID, out PatientInfo qPatient, out List<AssayInfo> Assay, out List<Order> qOrder);
                        var qAssay = from assay in Assay
                                     group assay by new
                                     {
                                         assay.Barcode,
                                         assay.BarcodeId,
                                         assay.BiomaterialCode
                                     } into gropedAssay
                                     select new
                                     {
                                         Barcode = gropedAssay.Key.Barcode,
                                         BarcodeId = gropedAssay.Key.BarcodeId,
                                         BiomaterialCode = gropedAssay.Key.BiomaterialCode
                                     };

                        if (qOrder.Count != 0)
                        {
                            //Создание тела XML сообщения для создания заказа в ЛИС 
                            XElement email;
                            XElement messagebody = new XElement("Request"
                                    , new XElement("RequestCode", qRequest.Code)
                                    , new XElement("AmbulantCard", qPatient.ID)
                                    , new XElement("HospitalCode", qRequest.FilialCode  == string.Empty ? Ini.HospitalCode : qRequest.FilialCode)
                                    , new XElement("HospitalName", qRequest.FilialLabel)
                                    , new XElement("DepartmentCode", qRequest.DepCode)
                                    , new XElement("DepartmentName", qRequest.DepLabel)
                                    , new XElement("PregnancyDuration", qPatient.SrokBeremennosti)
                                    , new XElement("CyclePeriod", qPatient.FazaCikla)
                                    , new XElement("Patient",
                                                   new XElement("LastName", qPatient.LastName),
                                                   new XElement("FirstName", qPatient.FirstName),
                                                   new XElement("MiddleName", qPatient.MiddleName),
                                                   new XElement("BirthDay", qPatient.BirthDate.Day),
                                                   new XElement("BirthMonth", qPatient.BirthDate.Month),
                                                   new XElement("BirthYear", qPatient.BirthDate.Year),
                                                   new XElement("Sex", qPatient.Gender == 0 ? 1 : qPatient.Gender == 1 ? 0 : 2),
                                                   email = new XElement("PrivatEmail", qPatient.Email)
                                                  )
                                    , new XElement("Samples",
                                                  from sample in qAssay
                                                  select new XElement("Sample",
                                                                new XElement("Barcode", sample.Barcode),
                                                                new XElement("BiomaterialCode", sample.BiomaterialCode),
                                                                new XElement("Targets",
                                                                from target in qOrder
                                                                where sample.BarcodeId == target.BarcodeId
                                                                select new XElement("Target",
                                                                             new XElement("Code", target.Code),
                                                                             new XElement("Priority", 0)
                                                                                    )
                                                                            )     //Targets
                                                                     )   //Sample
                                                    )   //Samples
                                    );    //Request
                            if (qPatient.Email == string.Empty) email.Remove();
                            messagebody.Save(string.Format("{0}/{1}.xml", Ini.OrderFilePath, qRequest.Code));
                            Console.WriteLine("Создан фaйл заказа {0}.xml и помещен в папку {1}", qRequest.Code, Ini.OrderFilePath);
                            DBControl.UpdateOrderState(qRequest.ID, 1);
                        }
                        else { Console.WriteLine("Файл заказа не создан, т.к. все услуги перенаправлены"); }
                        //Console.WriteLine("Barcode = {0} , BiomaterialCode = {1}", info.Barcode, info.BiomaterialCode);
                    }
                }
            }

            //Загрузка результатов в формате JSON
            static public void LoadResultsJson()
            {
                XDocument xdoc;
                int p_id = 0;

                string[] resfiles = Directory.GetFiles(Ini.ResultFilePath, "*.json");
                foreach (string filepath in resfiles)
                {
                    DataContext db = new DataContext(sqlConnectionStr);
                    Table<Impdata> TImpdata = db.GetTable<Impdata>();
                    /*
                    StreamReader rfile = new StreamReader(filepath);
                    js.JObject jdoc = js.JObject.Parse(rfile.ReadToEnd());
                    rfile.Close();

                    int name = (int)jdoc.SelectToken("OrderResponse.CardNumber");
                    Console.WriteLine(name);

                    /*
                    xdoc = XDocument.Load(filepath);
                    Console.WriteLine("Загрузка результатов из файла {0}", filepath);
                    string StrPatID = xdoc.Element("Request").Element("AmbulantCard").Value;

                    bool ok = true;

                    if (!DBControl.GetPatientID(db, StrPatID, filepath, out p_id))
                    {
                        ok = false;
                        PatientName Name = new PatientName
                        {
                            LastName = xdoc.Element("Request").Element("Patient").Element("LastName").Value,
                            FirstName = xdoc.Element("Request").Element("Patient").Element("FirstName").Value,
                            MiddleName = xdoc.Element("Request").Element("Patient").Element("MiddleName").Value
                        };

                        int BirthYear = 1990;
                        try { BirthYear = Convert.ToInt32(xdoc.Element("Request").Element("Patient").Element("BirthYear").Value); }
                        catch { }

                        PatientBirthDate BirthDate = new PatientBirthDate
                        {
                            BirthYear = BirthYear,
                            BirthDay = 0,
                            BirthMonth = 0
                        };

                        ok = DBControl.GetPatientID(db, Name, BirthDate, filepath, out p_id);
                    }

                    string KeyCode = xdoc.Element("Request").Element("OrderNumber").Value;

                    bool NewImp = DBControl.GetImpdataID(db, KeyCode, out int ImpdataId);

                    //Подсчет количесва параметров исследований
                    int resCount = (from xe in xdoc.Elements("Request").Elements("SampleResults").Elements("SampleResult").
                                       Elements("TargetResults").Elements("TargetResult").Elements("Works").Elements("Work")
                                    select xe.Element("Code")).Count();

                    DateTime resdate = Convert.ToDateTime(xdoc.Element("Request").Element("ResultDate").Value);

                    if (NewImp)
                    {
                        Console.WriteLine("Количетсво строк {0}, id записи {1}", resCount, ImpdataId);
                        var insetImpdata = from xe in xdoc.Elements("Request")
                                           select new Impdata
                                           {
                                               ImpdataId = ImpdataId,
                                               KEYCODE = xe.Element("OrderNumber").Value,
                                               Nom = xe.Element("Patient").Element("LastName").Value,
                                               Prenom = xe.Element("Patient").Element("FirstName").Value + " " + xe.Element("Patient").Element("MiddleName").Value,
                                               Date_Naissance = new DateTime(Convert.ToInt32(xe.Element("Patient").Element("BirthYear").Value), 1, 1),
                                               Date_Consultation = resdate,
                                               //PATIENTS_ID = p_id == 0 ? null : p_id,
                                               Mesure = Mesure
                                           };
                        //Вставка в Impdata
                        TImpdata.InsertAllOnSubmit(insetImpdata);
                        db.SubmitChanges();
                    }
                    //Выборка для вставки в DS_RESTESTS
                    List<RestestData> qRestests = (from xe in xdoc.Elements("Request").Elements("SampleResults").Elements("SampleResult").
                                                     Elements("TargetResults").Elements("TargetResult").Elements("Works").Elements("Work")
                                                   select new RestestData
                                                   {
                                                       VAL = xe.Element("Value").Value,
                                                       UNIT = xe.Element("UnitName").Value,
                                                       Rescode = xe.Element("Code").Value,
                                                       ResName = xe.Element("Name").Value,
                                                       NORM_TEXT_REC = xe.Element("Norm").Element("Norms").Value
                                                   }).ToList();

                    DBControl.InsertRestestData(db, qRestests, ImpdataId, resdate);

                    //Выборка и вставка комменариев
                    var Comments = xdoc.Element("Request").Element("SampleResults").Element("SampleResult").
                                  Element("TargetResults").Element("TargetResult");

                    if (ok) DBControl.UpdateImpdata(db, ImpdataId, p_id);
                    db.Dispose();

                    if (ok)
                    {
                        DBControl.ErrInsertToMotconsu(ImpdataId, p_id, resdate);
                        DBControl.InsetToUserTables(ImpdataId);
                    }

                    if (Directory.Exists(Ini.ArchiveFilePath))
                    {
                        string fp = string.Format("{0}/{1}", Ini.ArchiveFilePath, Path.GetFileName(filepath));
                        if (File.Exists(fp)) File.Delete(fp);
                        File.Copy(filepath, fp);
                        File.Delete(filepath);
                    }
                    else
                        File.Delete(filepath);
                    //*/
                }
            }


            //Формирование и выгрузка заказов в формате JSON
            static public void UploadOrdersJson()
            {
                CultureInfo culture = new CultureInfo("ru-RU"); //формат преобразования даты
                List<Request> RequestInfoList = new List<Request>();

                DBControl.GetOrdersList(0, 5, out RequestInfoList);

                if (RequestInfoList.Count == 0) Console.WriteLine("Нет направлений для заказов");
                else Console.WriteLine("Готово {0} заказов", RequestInfoList.Count);

                foreach (Request qRequest in RequestInfoList)
                {
                    //Получение данных заказа из базы данных
                    DBControl.GetOrderInfoV4(qRequest.ID, out PatientInfo qPatient, out List<AssayInfo> Assay, out List<Order> qOrder);

                    var qAssay = from assay in Assay
                                 group assay by new
                                 {
                                     assay.Barcode,
                                     assay.BarcodeId,
                                     assay.BiomaterialCode
                                 } into gropedAssay
                                 select new
                                 {
                                     Barcode = gropedAssay.Key.Barcode,
                                     BarcodeId = gropedAssay.Key.BarcodeId,
                                     BiomaterialCode = gropedAssay.Key.BiomaterialCode
                                 };

                    if (qOrder.Count != 0)
                    {
                        //Создание тела XML сообщения для создания заказа в ЛИС 
                        /*
                        js.JObject messagebody = new js.JObject(
                                new js.JProperty("OrderRequest", new js.JObject(
                                       new js.JProperty("RequestCode", qRequest.Code)
                                     , new js.JProperty("AmbulantCard", qPatient.ID)
                                     , new js.JProperty("HospitalCode", qRequest.FilialCode)
                                     , new js.JProperty("HospitalName", qRequest.FilialLabel)
                                     , new js.JProperty("DepartmentCode", qRequest.DepCode)
                                     , new js.JProperty("DepartmentName", qRequest.DepLabel)
                                     , new js.JProperty("Patient", new js.JObject(
                                                new js.JProperty("LastName", qPatient.LastName)
                                              , new js.JProperty("FirstName", qPatient.FirstName)
                                              , new js.JProperty("MiddleName", qPatient.MiddleName)
                                              , new js.JProperty("BirthDay", qPatient.BirthDate.Day)
                                              , new js.JProperty("BirthMonth", qPatient.BirthDate.Month)
                                              , new js.JProperty("BirthYear", qPatient.BirthDate.Year)
                                              , new js.JProperty("Sex", qPatient.Gender == 0 ? 1 : qPatient.Gender == 1 ? 2 : 0)
                                            ))
                                     , new js.JProperty("Samples", new js.JObject(
                                                new js.JProperty("Sample", new js.JArray(
                                                    from sample in qAssay
                                                    select new js.JObject(
                                                        new js.JProperty("Barcode", sample.Barcode)
                                                      , new js.JProperty("BiomaterialCode", sample.BiomaterialCode) //
                                                      , new js.JProperty("Targets", new js.JObject(
                                                            new js.JProperty("Target", new js.JArray(
                                                                from target in qOrder
                                                                where sample.BarcodeId == target.BarcodeId
                                                                select new js.JObject(
                                                                      new js.JProperty("Code", target.Code)
                                                                    , new js.JProperty("Priority", 0)
                                                              ) //Target
                                                        )) //Target Array
                                                    )) //Targets
                                              ) //Sample 
                                         )) //Sample Array
                                     )) //Samples
                                 )) //OrderRequest
                             ); //messagebody
                             
                        /*         
                                 , new js.JProperty("AmbulantCard", qPatient.ID)
                                 , new js.JProperty("HospitalCode", qRequest.FilialCode)
                                 , new js.JProperty("HospitalName", qRequest.FilialLabel)
                                 , new js.JProperty("DepartmentCode", qRequest.DepCode)
                                 , new js.JProperty("DepartmentName", qRequest.DepLabel)
                                 , new js.JProperty("Patient",
                                                new js.JProperty("LastName", qPatient.LastName),
                                                new js.JProperty("FirstName", qPatient.FirstName),
                                                new js.JProperty("MiddleName", qPatient.MiddleName),
                                                new js.JProperty("BirthDay", qPatient.BirthDate.Day),
                                                new js.JProperty("BirthMonth", qPatient.BirthDate.Month),
                                                new js.JProperty("BirthYear", qPatient.BirthDate.Year),
                                                new js.JProperty("Sex", qPatient.Gender == 0 ? 1 : qPatient.Gender == 1 ? 2 : 0)
                                               )
                                 , new js.JProperty("Samples",
                                               from sample in qAssay
                                               select new js.JProperty("Sample",
                                                             new js.JProperty("Barcode", sample.Barcode),
                                                             new js.JProperty("BiomaterialCode", sample.BiomaterialCode),
                                                             new js.JProperty("Targets",
                                                             from target in qOrder
                                                             where sample.BarcodeId == target.BarcodeId
                                                             select new js.JProperty("Target",
                                                                          new js.JProperty("Code", target.Code),
                                                                          new js.JProperty("Priority", 0)
                                                                                 )
                                                                         )     //Targets
                                                                  )   //Sample
                                                 )   //Samples
                                 );    //Request
                         //*/
                        //messagebody.Save(string.Format("{0}/{1}.xml", Ini.OrderFilePath, qRequest.Code));
                        StreamWriter wfile = new StreamWriter(@"D:\data\1234.json", false);
                        //wfile.Write(messagebody.ToString());
                        wfile.Close();

                        Console.WriteLine("Создан фaйл заказа {0}.xml и помещен в папку {1}", qRequest.Code, Ini.OrderFilePath);
                        //DBControl.UpdateOrderState(qRequest.ID, 1);
                    }
                    else { Console.WriteLine("Файл заказа не создан, т.к. все услуги перенаправлены"); }
                    //Console.WriteLine("Barcode = {0} , BiomaterialCode = {1}", info.Barcode, info.BiomaterialCode);
                }

            }

        }

        class SLS
        {
            static bool RequestFiles(string miscode, string misWPass, out string[] FileNames)
            {
                bool result = false;
                FileNames = null;
                string messagebody = Ini.url + "/" + "download_files.pl" //выгрузка файлов
                                    + "?" + "username=" + miscode + "&" + "pass=" + misWPass   // логин и пароль
                                    + "&" + "directory=..%2Fobmen%2F" + miscode + "%2Fresult" // директория расположения файлов для загрузки
                                    + "&" + "filename=" + "all"; //выгрузка названий файлов
                                    
                Message.SLS filesInfo = new Message.SLS();

                if (filesInfo.RequestGET(messagebody, out XDocument xdoc, out string errmess))
                {
                    if (Ini.WriteBuffer) Files.ToWriteFile(xdoc.ToString(), 0);

                    XNamespace xmlns = "http://www.w3.org/1999/xhtml";
                    XElement titlmessElement = xdoc.Element(xmlns + "html").Element(xmlns + "head").Element(xmlns + "title");
                    string titlmess = "";
                    if (titlmessElement != null) titlmess = titlmessElement.Value;

                    if (titlmess == "Files for download")
                    {
                        FileNames = (xdoc.Element(xmlns + "html").Element(xmlns + "body").Element(xmlns + "p").Value).Split(' ');
                        result = true;
                    }
                    else
                    {
                        Console.WriteLine(titlmess);
                        Files.ToWriteFile(titlmess, 2);
                    }
                }
                else if (errmess != "") Files.ToWriteFile(errmess, 2);
                return result;
            }

            static bool RequestResults(string miscode, string misWPass, string FileName, out XDocument xdoc)
            {
                string messagebody = Ini.url + "/" + "download_files.pl" //выгрузка файлов
                                    + "?" + "username=" + miscode + "&" + "pass=" + misWPass   // логин и пароль
                                    + "&" + "directory=..%2Fobmen%2F" + miscode + "%2Fresult" // директория расположения файлов для загрузки
                                    + "&" + "filename=" + FileName; //выгрузка результатов 
 
                Message.SLS filesInfo = new Message.SLS();

                bool res = filesInfo.RequestGET(messagebody, out xdoc, out string errmess);
                if (errmess != "") Files.ToWriteFile(errmess, 2);
                return res;
            }

            static bool RequestResultsDelete(string miscode, string misWPass, string FileName)
            {
                string messagebody = Ini.url + "/" + "download_files.pl" //выгрузка файлов
                                    + "?" + "username=" + miscode + "&" + "pass=" + misWPass   // логин и пароль
                                    + "&" + "directory=..%2Fobmen%2F" + miscode + "%2Fresult" // директория расположения файлов для загрузки
                                    + "&" + "filename=" + FileName //выгрузка результатов 
                                    + "&" + "with_delete=1"; //удаление файла
                XDocument xdoc;
                Message.SLS filesInfo = new Message.SLS();

                bool res = filesInfo.RequestGET(messagebody, out xdoc, out string errmess);
                if (errmess != "") Files.ToWriteFile(errmess, 2);
                return res;
            }

            static public void LoadResults()
            {
                if (Ini.TestMode == 0)
                    foreach (var filial in FilialInfoList)
                    {
                        if (RequestFiles(filial.Code, filial.Pass, out string[] FileNames))
                        {
                            foreach (string fName in FileNames)
                            {
                                try
                                {
                                    if (RequestResults(filial.Code, filial.Pass, fName, out XDocument xdoc))
                                    {
                                        InsertXmlData(fName, xdoc);
                                        if (Ini.DeleteDownloaded)
                                            if (RequestResultsDelete(filial.Code, filial.Pass, fName))
                                                Console.WriteLine("Файл {0} удаелн с сервера", fName);
                                    }
                                }
                                catch(Exception ex)
                                {
                                    Files.WriteLogFile(string.Format("Ошибка обработки файала {0}: ", fName, ex.ToString()));
                                    Console.WriteLine(ex.Message);
                                }
                            }
                        }
                    }
                else
                {
                    if (Ini.TestMode == 1)
                    {
                        XDocument xdoc;
                        string fileName;
                        string[] resfiles = Directory.GetFiles(Ini.ResultFilePath, "*.xml");
                        foreach (string filepath in resfiles)
                        {
                            xdoc = XDocument.Load(filepath);
                            fileName = Path.GetFileName(filepath);
                            InsertXmlData(fileName, xdoc);
                        }
                    }

                }

            }

            static public void GetResultValues(List<RestestData> qRestests, out List<RestestData> resultList, out List<RestestData> antibioticList)
            {
                resultList = new List<RestestData>();
                antibioticList = new List<RestestData>();
                RestestData antibiotic;
                RestestData result;
                int num = 0;

                foreach (var item in qRestests)
                {
                    string[] resultArray = item.VAL.Split(new char[] { '&' }, StringSplitOptions.RemoveEmptyEntries);
                    string comments = string.Empty;
                    int i = 0;
                    if (resultArray.Count() > 1)
                    {
                        foreach (string resStr in resultArray)
                        {
                            Console.WriteLine(resStr);
                            
                            if (resStr.IndexOf("1A") != -1)
                            {
                                num++;
                                result = new RestestData
                                {
                                    VAL = resStr.Replace("1A","").Trim(new char[] { '|'}),
                                    UNIT = item.UNIT,
                                    Rescode = item.Rescode,
                                    NORM_TEXT_REC = item.NORM_TEXT_REC.Replace("&", "\n"),
                                    Item = num
                                };
                                resultList.Add(result);
                            }
                            if (resStr.IndexOf("2B") != -1)
                            {
                                string[] antibioticArray = resStr.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                                antibiotic = new RestestData
                                {
                                    Rescode = antibioticArray[0].Replace("2B", ""),
                                    ResName = antibioticArray[0].Replace("2B", ""),
                                    VAL = antibioticArray[1],
                                    DataType = 8,
                                    Item = num
                                };
                                antibioticList.Add(antibiotic);
                            }

                            if (resStr.IndexOf("1B") != -1)
                            {
                                string r = CheckSpaces(resStr);
                                comments += i==0 ? r.Replace("1B", "").Trim(new char[] { '|' }) : " \n" + r.Replace("1B", "").Trim(new char[] {'|' });
                                i++;
                            }

                        }
                    }
                    else
                    {
                        result = new RestestData
                        {
                            VAL = item.VAL.Replace("1B", "").Trim(new char[] { '|' }),
                            UNIT = item.UNIT,
                            Rescode = item.Rescode,
                            NORM_TEXT_REC = item.NORM_TEXT_REC.Replace("&", "\n"),
                        };
                        resultList.Add(result);
                    }

                    if (comments != string.Empty)
                    {
                        result = new RestestData
                        {
                            VAL = comments,
                            Rescode = "AnaComment",
                            ResName = "Комментарий"
                        };
                        resultList.Add(result);
                    }

                }
            }

            static public void InsertXmlData(string fName, XDocument xdoc)
            {
                Console.WriteLine("Загружен файл результатов {0}, обработка файла... ", fName);
                Files.ToWriteFile(string.Format("Загружен файл результатов {0}", fName), 2);
                if (Ini.WriteInput)
                {
                    xdoc.Save(string.Format("{0}/{1}", Files.InputDir, fName));
                }

                string IsDone = xdoc.Element("root").Element("Order").Element("Tests").Element("Test1").Attribute("IsDone") == null ? "" :
                     xdoc.Element("root").Element("Order").Element("Tests").Element("Test1").Attribute("IsDone").Value;

                if (IsDone == "1")
                {
                    DataContext db = new DataContext(sqlConnectionStr);
                    Table<Impdata> TImpdata = db.GetTable<Impdata>();

                    XElement PatInfo = xdoc.Element("root").Element("Order").Element("Patient");
                    bool ok = false;
                    string StrPatID;

                    if (PatInfo.Element("MedCardID") != null) StrPatID = PatInfo.Element("MedCardID").Value;
                    else StrPatID = "";

                    string RecordId = xdoc.Element("root").Element("Order").Attribute("OrderID").Value;

                    string[] NameArray = PatInfo.Element("LastName").Value.Split(' ');
                    string[] IOArray;
                    string IO = NameArray.Count() > 1 ? NameArray[1] : "";
                    string FName = "";
                    string MName = "";

                    if (IO.IndexOf('.') > -1 && NameArray.Count() < 3)
                    {
                        IOArray = IO.Split('.');
                        FName = IOArray[0];
                        MName = IOArray.Count() > 1 ? IOArray[1] : "";
                    }

                    PatientName Name = new PatientName
                    {
                        LastName = NameArray[0],
                        FirstName = FName == "" ? (NameArray.Count() > 1 ? NameArray[1] : "") : FName,
                        MiddleName = MName == "" ? (NameArray.Count() > 2 ? NameArray[2] : "") : MName,
                    };

                    if (!DateTime.TryParse(PatInfo.Element("DOB").Value, out DateTime PatBirthDate))
                    {
                        if (PatInfo.Element("DOB").Value.Length == 4 & int.TryParse(PatInfo.Element("DOB").Value, out int BYear))
                            PatBirthDate = new DateTime(BYear, 1, 1);
                        else PatBirthDate = new DateTime(1900, 1, 1);
                    }
                    ok = DBControl.GetPatientID(db, StrPatID, RecordId, out int p_id);
                    if (!ok)
                    {
                        PatientBirthDate BirthDate = new PatientBirthDate
                        {
                            BirthYear = PatBirthDate.Year,
                            BirthMonth = PatBirthDate.Month,
                            BirthDay = PatBirthDate.Day
                        };
                        ok = DBControl.GetPatientID(db, Name, BirthDate, RecordId, out p_id);
                    }

                    if (Ini.TestPatientID != 0) { ok = true; p_id = Ini.TestPatientID; }

                    string KeyCode = xdoc.Element("root").Element("Order").Attribute("OrderID").Value;

                    bool NewImp = DBControl.GetImpdataID(db, KeyCode, out int ImpdataId);

                    DateTime.TryParseExact(string.Format("{0} {1}", xdoc.Element("root").Element("Header").Element("FileDate").Value,
                              xdoc.Element("root").Element("Header").Element("FileTime").Value),
                              new string[] { "yyMMdd HH:mm:ss", "yyMMdd h:mm:ss" }, null, DateTimeStyles.None, out DateTime resdate);

                    if (NewImp)
                    {
                        Console.WriteLine("Создана запиь импорта, id записи {0}", ImpdataId);

                        var insetImpdata = from xe in xdoc.Elements("root")
                                           select new Impdata
                                           {
                                               ImpdataId = ImpdataId,
                                               KEYCODE = KeyCode,
                                               Nom = Name.LastName,
                                               Prenom = string.Format("{0} {1}", Name.FirstName, Name.MiddleName),
                                               Date_Naissance = PatBirthDate,
                                               Date_Consultation = resdate,
                                               Mesure = Mesure,
                                               FilialCode = xe.Element("Header").Element("ClinicID").Value
                                           };
                        //Вставка в Impdata
                        TImpdata.InsertAllOnSubmit(insetImpdata);
                        db.SubmitChanges();
                    }
                    //System.Xml.Linq.Extensions.
                    //Выборка для вставки в DS_RESTESTS

                    List<RestestData> qRestests = (from xe in xdoc.Elements("root").Elements("Order").Elements("Tests").Descendants().Elements("Result")
                                                   where xe.Element("res") != null
                                                   select new RestestData
                                                   {
                                                       VAL = xe.Element("res").Value.Trim(new char[] { '|', '&' }),
                                                       UNIT = xe.Element("quant") == null ? "" : xe.Element("quant").Value,
                                                       Rescode = xe.Attribute("id").Value,
                                                       ResName = "",
                                                       NORM_TEXT_REC = xe.Element("norm") == null ? "" : xe.Element("norm").Value
                                                   }).ToList();

                    GetResultValues(qRestests, out List<RestestData> results, out List<RestestData> antibiotics);

                    DBControl.InsertRestestData(db, results, ImpdataId, resdate);

                    //Выборка и вставка антибоиотиков
                    if (antibiotics.Count() > 0)
                    {
                        DBControl.InsertRestestData(db, antibiotics, ImpdataId, resdate);

                        //Построение лабораторных справочников: Чувствительность к антибиотикам
                        if (Ini.AntbGroupId != 0)
                        {
                            DBControl.BuildLabDitionary(db, Ini.AntbGroupId, 8);
                        }
                    }

                    if (ok & NewImp) DBControl.UpdateImpdata(db, ImpdataId, p_id);

                    db.Dispose();

                    if (ok)
                    {
                        DBControl.ErrInsertToMotconsu(ImpdataId, p_id, resdate, NewImp);
                        //DBControl.InsetToUserTables(ImpdataId);
                        //Дублирование комментариев в отдельную таблицу
                        DBControl.InsertToEMCTables(ImpdataId);
                        DBControl.UpdateDirAnsw(ImpdataId);
                    }
                }

            }

            static public string  CheckSpaces(string s)
            {
                if (s.IndexOf("   ") > 0)
                {
                    int i = 1;

                    int len = s.Length;
                    while (len > 3)
                    {
                        if (s[i-1] == ' ' & s[i] == ' ' & s[i + 1] == ' ')
                        {
                            s = s.Remove(i + 1, 1);
                        }
                        else i++;
                        len--;
                      }
                }
                return s;
            }

            static public void UploadOrders()
            {
                CultureInfo culture = new CultureInfo("ru-RU"); //формат преобразования даты
                List<Request> RequestInfoList = new List<Request>();

                //DBControl.GetOrdersList(out RequestInfoList);
                DBControl.GetOrdersList(0, 5, out RequestInfoList);

                if (RequestInfoList.Count == 0) Console.WriteLine("Нет направлений для заказов");
                else Console.WriteLine("Готово {0} заказов", RequestInfoList.Count);

                ///*
                if (RequestInfoList.Count != 0) 
                {
                    Console.WriteLine("Отправка готовых заказов...");

                    foreach (Request qRequest in RequestInfoList)
                    {
                        //Получение данных заказа из базы данных
                        //DBControl.GetOrderInfoV3(qRequest.ID, out PatientInfo qPatient, out List<AssayInfo>  Assay, out List<Order> qOrder);  //Триггер по типу Астромеда
                        DBControl.GetOrderInfoV4_1(qRequest.ID, out PatientInfo qPatient, out List<AssayInfo> Assay, out List<Order> qOrder);   //Триггер по типу СОМЦ (разделение направлений на разные заказы)
                        var qAssay = from assay in Assay
                                     group assay by new
                                     {
                                         assay.Barcode,
                                         assay.BarcodeId
                                     } into gropedAssay
                                     select new
                                     {
                                         Barcode = gropedAssay.Key.Barcode,
                                         BarcodeId = gropedAssay.Key.BarcodeId
                                     };

                        //Создание тела XML сообщения для создания заказа в ЛИС 
                        string messageId = (from m in qAssay
                                            orderby m.Barcode
                                            select m.Barcode
                                            ).First();

                        messageId = messageId.Length == 10 ? qRequest.WebId + messageId : messageId;

                        XElement messagebody = new XElement("root"
                            , new XElement("Header"
                                , new XElement("Version", "5")
                                , new XElement("FileDate", qRequest.Date.ToString("yyyy-MM-dd"))
                                , new XElement("FileTime", qRequest.Date.ToString("HH:mm:ss"))
                                , new XElement("FileType", "Order")
                                , new XElement("ClinicID", qRequest.FilialCode)
                                , new XElement("ClinicName", qRequest.FilialLabel)
                                , new XElement("Direction", "FromOrganization")
                                            )
                            , new XElement("Order", new XAttribute("OrderID", messageId)
                                    , new XElement("Patient", new XAttribute("PatientID", messageId),
                                             new XElement("MedCardID", qPatient.ID),
                                             new XElement("LastName", string.Format("{0} {1} {2}", qPatient.LastName, qPatient.FirstName, qPatient.MiddleName)),
                                             new XElement("DOB", qPatient.BirthDate.ToString("yyyy-MM-dd")),
                                             new XElement("Sex", qPatient.Gender == 0 ? "M" : qPatient.Gender == 1 ? "Ж" : "Ж")
                                                    )
                                    , new XElement("Tests",
                                              from test in qOrder
                                              select new XElement("Test", new XAttribute("TestID", test.Code),
                                                            new XElement("TestName", test.Name),
                                                            from sample in qAssay
                                                            where sample.BarcodeId == test.BarcodeId
                                                            select new XElement("BarCode", sample.Barcode)
                                                                   )     //Test
                                                    )   //Tests
                                            )   //Order
                                ); //root

                        string filename = string.Format("{0}.xml", messageId);  //qRequest.Code

                        if (Ini.WriteOutput) messagebody.Save(string.Format("{0}/{1}", Files.OutputDir, filename));

                        XDocument xdoc = new XDocument(new XDeclaration("1.0", "windows-1251",""), messagebody);

                        ///*
                        if (Ini.TestMode != 2)
                        {
                            Message.SLS sls = new Message.SLS();
                            if (Ini.TestMode == 1)
                            {
                                qRequest.FilialCode = Ini.mis;
                                qRequest.FilialPass = Ini.WPass;
                            }

                            if (sls.RequestPost(xdoc, filename, qRequest.FilialCode, qRequest.FilialPass, out XDocument response, out string errmessage))
                            {
                                Console.WriteLine("Ответ вэб-сервиса: \r\n {0}", response);

                                XNamespace xmlns = "http://www.w3.org/1999/xhtml";
                                if (Ini.WriteBuffer) Files.ToWriteFile(response.ToString(), 1);
                                string respHeder = "";
                                string resp = "";
                                XElement respHeaderElement = response.Element(xmlns + "html").Element(xmlns + "head").Element(xmlns + "title");
                                if (respHeaderElement != null) respHeder = respHeaderElement.Value;
                                XElement respElement = response.Element(xmlns + "html").Element(xmlns + "body").Element(xmlns + "p");
                                if (respElement != null) resp = respElement.Value;

                                if (respHeder == "Error")
                                {
                                    Files.ToWriteFile(string.Format("{0} Файл: {1}", resp, filename), 2);
                                }
                                if (respHeder == "Success")
                                {
                                    Files.ToWriteFile(string.Format("Выгружен файл {0}", filename), 2);
                                    //Изменение статуса заказа
                                    DBControl.UpdateOrderInfo(qRequest.ID, 2, messageId);
                                }
                            }
                            else
                            {
                                if (errmessage != "") Files.ToWriteFile(string.Format("{0} Файл: {1}", errmessage, filename), 2); 
                            }
                        }

                    }
                }
                //*/
            }

        }
   
        class INVITRO
        {
            //Загрузка результатов
            static public void LoadResults()
            {
                //if (!Files.GoFurther()) goto EndOfWork;
                XDocument xdoc = null;
                int p_id = 0;
                string filialAddressCode;

                string[] resfiles = Directory.GetFiles(Ini.ResultFilePath, "*.xml");
                bool attachPdf = Directory.Exists(Ini.PdfFilePath) && (Ini.RubricsId != 0);
                int motconsuId = 0;
                DateTime resdate = DateTime.Now;
                foreach (string filepath in resfiles)
                {
                    DataContext db = new DataContext(sqlConnectionStr);
                    Table<Impdata> TImpdata = db.GetTable<Impdata>();

                    xdoc = XDocument.Load(filepath);
                    Console.WriteLine("Загрузка результатов из файла {0}", filepath);

                    if (xdoc.Element("SafirMessage") == null)
                    {
                        File.Delete(filepath);
                        goto Finish;
                    }

                    if (xdoc.Element("SafirMessage").Element("Envelope").Element("Requisition") != null) Ini.Option = 0;
                    else Ini.Option = 1;
                    XElement Patient = Ini.Option == 0 ? xdoc.Element("SafirMessage").Element("Envelope").Element("Requisition").Element("Patient")
                                                        : xdoc.Element("SafirMessage").Element("Requisition").Element("Patient");

                    string[] FirstName = Patient.Attribute("FirstName").Value.Split(' ');

                    bool ok = true;

                    PatientName Name = new PatientName
                    {
                        LastName = Patient.Attribute("Surname").Value,
                        FirstName = FirstName[0],
                        MiddleName = FirstName.Count() > 1 ? FirstName[1] : string.Empty
                    };

                    DateTime PatBirthDate = DateTime.ParseExact(Patient.Attribute("BirthDate").Value, "yyyyMMdd", null);

                    PatientBirthDate BirthDate = new PatientBirthDate
                    {
                        BirthYear = PatBirthDate.Year,
                        BirthMonth = PatBirthDate.Month,
                        BirthDay = PatBirthDate.Day
                    };

                    ok = DBControl.GetPatientID(db, Name, BirthDate, filepath, out p_id);

                    //Переопределение ID отделения и MedDepId по коду в XML
                    XElement ReqUnit = Ini.Option == 0 ? xdoc.Element("SafirMessage").Element("Envelope").Element("Requisition").Element("ReqUnit")
                                       : xdoc.Element("SafirMessage").Element("Requisition").Element("ReqUnit");
                    filialAddressCode = ReqUnit.Attribute("AddressCode").Value;
                    Console.WriteLine("Code  {0}, MedId {1}", filialAddressCode, MedicinsId);

                    //Опция для нескольких филиалов
                    /*
                      DBControl.GetDepId(filialAddressCode, Ini.MedecinId == 0 ? MedicinsId : Ini.MedecinId, out int depId, out int medDepId);
                      MedDepId = medDepId; DepId = depId;
                    Console.WriteLine("Meddep {0} Dep {1}", MedDepId, DepId);
                    */
                    //
                    XElement ResultsData = Ini.Option == 0 ? xdoc.Element("SafirMessage").Element("Envelope").Element("Requisition").Element("Reply")
                                                           : xdoc.Element("SafirMessage").Element("Requisition").Element("Reply");

                    string KeyCode = xdoc.Element("SafirMessage").Element("Envelope").Element("Message").Attribute("ID").Value;

                    bool NewImport = DBControl.GetImpdataID(db, KeyCode, out int ImpdataId);

                    //Подсчет количества параметров исследований  Element("Analysis")
                    int resCount = (from xe in ResultsData.Elements("Sample").Elements("Analysis")
                                    select xe.Attribute("TestMethodCode")).Count();

                    Console.WriteLine("Количетсво строк {0}, id записи {1}", resCount, ImpdataId);

                    resdate = DateTime.ParseExact(xdoc.Element("SafirMessage").Element("Envelope").Element("Sent").Attribute("DateTime").Value
                                    , "yyyyMMddHHmmss", null);

                    //Вставка в Impdata
                    if (NewImport)
                    {
                        ImpData impData = new ImpData()
                        {
                            ImpdataId = ImpdataId,
                            KEYCODE = KeyCode,
                            Nom = Name.LastName,
                            Prenom = Name.FirstName + " " + Name.MiddleName,
                            Date_Naissance = new DateTime(BirthDate.BirthYear, BirthDate.BirthMonth, BirthDate.BirthDay),
                            Date_Consultation = resdate,
                            Mesure = Mesure
                        };

                        DBControl.InsertImpdata(db, impData);
                    }

                    RestestData restestData;

                    XElement comments = Ini.Option == 0 ? xdoc.Element("SafirMessage").Element("Envelope").Element("Requisition").Element("ReqComment")
                                                        : xdoc.Element("SafirMessage").Element("Requisition").Element("ReqComment");
                    //Комментарии к заявке
                    if (comments != null)
                    {
                        var rComment = from xe in comments.Attributes("Text")
                                       select xe.Value;

                        string ReqComment = string.Empty;
                        int i = 0;
                        foreach (var item in rComment)
                        {
                            ReqComment += i == 0 ? item : Convert.ToString(Convert.ToChar(13)) + item;
                            i++;
                        }

                        restestData = new RestestData
                        {
                            M_VAL = ReqComment,
                            Rescode = "ReqComment"
                        };

                        DBControl.InsertRestestItem(db, restestData, ImpdataId, resdate);

                    };

                    //Название исследования

                    restestData = (from xe in ResultsData.Elements("Sample").Elements("Analysis")
                                   where xe.Attribute("Value").Value == string.Empty
                                   select new RestestData
                                   {
                                       VAL = xe.Attribute("AnaName").Value,
                                       Rescode = "AnaName"
                                   }).FirstOrDefault();

                    if (restestData != null)
                        DBControl.InsertRestestItem(db, restestData, ImpdataId, resdate);

                    //Комментарии к исследованию
                    if (ResultsData.Element("Sample").Element("Analysis") != null)
                    {
                        if (ResultsData.Element("Sample").Element("Analysis").Element("AnaComment") != null)
                        {
                            restestData = (from xe in ResultsData.Elements("Sample").Elements("Analysis")
                                           where xe.Attribute("Value").Value == string.Empty
                                           select new RestestData
                                           {
                                               M_VAL = xe.Element("AnaComment").Attribute("Text").Value,
                                               Rescode = "AnaComment"
                                           }).FirstOrDefault();

                            if (restestData != null)
                                DBControl.InsertRestestItem(db, restestData, ImpdataId, resdate);
                        }
                    }

                    //Выборка для вставки в DS_RESTESTS
                    List<RestestData> qRestests = (from xe in ResultsData.Elements("Sample").Elements("Analysis")
                                                   where xe.Attribute("Value").Value != string.Empty
                                                   select new RestestData
                                                   {
                                                       VAL = xe.Attribute("Value").Value,
                                                       UNIT = xe.Attribute("Unit") == null ? "" : xe.Attribute("Unit").Value,
                                                       Rescode = xe.Attribute("TestMethodCode").Value,
                                                       ResName = xe.Attribute("AnaName").Value,
                                                       NORM_TEXT_REC = xe.Attribute("RefText") != null ? xe.Attribute("RefText").Value :
                                                           (xe.Attribute("RefMin") != null & xe.Attribute("RefMax") != null) ?
                                                           xe.Attribute("RefMin").Value + " - " + xe.Attribute("RefMax").Value : "",
                                                       Comments = xe.Element("AnaComment") != null ? xe.Element("AnaComment").Attribute("Text").Value : ""
                                                   }).ToList();

                    DBControl.InsertRestestData(db, qRestests, ImpdataId, resdate);

                    //Построение лабораторных справочников
                    if (Ini.ResGroupId != 0)
                    {
                        DBControl.BuildLabDitionary(db, Ini.ResGroupId, 0);
                    }

                    //Выборка и вставка микроорганизмов

                    if (ResultsData.Elements("Sample").Elements("Analysis").Elements("Culture").Count() > 0)
                    {
                        int i = 1;
                        var culture = (from xe in ResultsData.Elements("Sample").Elements("Analysis").Elements("Culture")
                                       select new Culture
                                       {
                                           Number = i++,
                                           Label = xe.Attribute("Finding").Value,
                                           Growth = xe.Attribute("Growth") == null ? "-" : xe.Attribute("Growth").Value
                                       }).ToList();

                        qRestests = (from cl in culture
                                     select new RestestData
                                     {
                                         VAL = cl.Growth,
                                         Rescode = cl.Label,
                                         ResName = cl.Label,
                                         Item = cl.Number,
                                         DataType = 7
                                     }
                                     ).ToList();

                        DBControl.InsertRestestData(db, qRestests, ImpdataId, resdate);

                        //Построение лабораторных справочников: Выявленные мироорганизмы
                        if (Ini.CultGroupId != 0)
                        {
                            DBControl.BuildLabDitionary(db, Ini.CultGroupId, 7);
                        }

                        //Выборка и вставка антибоиотиков
                        qRestests = (from xe in ResultsData.Elements("Sample").Elements("Analysis").Elements("Culture").Elements("Resistence")
                                     join cl in culture on xe.Parent.Attribute("Finding").Value equals cl.Label
                                     select new RestestData
                                     {
                                         VAL = xe.Attribute("SIR").Value,
                                         Rescode = xe.Attribute("Antibiotics").Value,
                                         ResName = xe.Attribute("Antibiotics").Value,
                                         Item = cl.Number,
                                         DataType = 8
                                     }
                                    ).ToList();

                        DBControl.InsertRestestData(db, qRestests, ImpdataId, resdate);

                        //Построение лабораторных справочников: Чувствительность к антибиотикам
                        if (Ini.AntbGroupId != 0)
                        {
                            DBControl.BuildLabDitionary(db, Ini.AntbGroupId, 8);
                        }
                    }

                    //Отправка результатов на тестового пациента
                    if (Ini.TestPatientID != 0) { ok = true; p_id = Ini.TestPatientID; }

                    string fileNamePart = Path.GetFileName(filepath).Split(new char[] { '_' })[0];
                    string[] pdffilepathes = Directory.GetFiles(Ini.ResultFilePath, fileNamePart + "*.pdf");

                    //Создание записи в ЭМК
                    if (ok)
                    {
                        DBControl.UpdateImpdata(db, ImpdataId, p_id);
                        //Запись в ЭМК
                        motconsuId = DBControl.InsertToEMCe(ImpdataId, p_id, resdate, NewImport);

                        //Дублирование комментариев в отдельную таблицу
                        //DBControl.InsertToEMCTables(ImpdataId);  

                        //Дублирование длинных комментариев к параметрам
                        if (Ini.CommentTable != null && Ini.CommentField != null)
                            DBControl.ParamsComments(ImpdataId, motconsuId, p_id, Ini.CommentTable, Ini.CommentField);

                        //Обработка pdf файла и прикрепление к ЭМК
                        if (attachPdf && motconsuId != 0)
                        {
                            string folder = string.Format("{0}-{1}/{2}", resdate.Year, Convert.ToString(resdate.Month).PadLeft(2, '0'), p_id);
                            string pdfDirectory = string.Format("{0}/{1}", Ini.PdfFilePath, folder);

                            if (!Directory.Exists(pdfDirectory))
                            {
                                DirectoryInfo dir = Directory.CreateDirectory(pdfDirectory);
                            }

                            foreach (string pdffp in pdffilepathes)
                            {
                                string medialogfp = string.Format("{0}/{1}", pdfDirectory, Path.GetFileName(pdffp));
                                if (File.Exists(medialogfp)) File.Delete(medialogfp);
                                if (File.Exists(pdffp))
                                {
                                    File.Copy(pdffp, medialogfp);
                                    File.Delete(pdffp);
                                    DBControl.AttachPdf(motconsuId, p_id, Path.GetFileName(pdffp), resdate, folder, Processing.Ini.RubricsId);
                                }
                            }
                        }
                    }

                    db.Dispose();
                    ///*
                    //Отправка файла во временную директорию или удаление
                    if (Directory.Exists(Ini.ArchiveFilePath))
                    {
                        string archivefp = string.Format("{0}/{1}", Ini.ArchiveFilePath, Path.GetFileName(filepath));
                        if (File.Exists(archivefp)) File.Delete(archivefp);
                        File.Copy(filepath, archivefp);
                        File.Delete(filepath);
                    }
                    else
                    {
                        if (Ini.DeleteDownloaded)
                        {
                            File.Delete(filepath);
                        }
                    }

                    if (Directory.Exists(Ini.UnboundPdfFilePath))
                    {
                        string unbound = string.Format("{0}/{1}", Ini.UnboundPdfFilePath, Path.GetFileName(filepath));
                        foreach (string pdffp in pdffilepathes)
                        {
                            if (File.Exists(pdffp))
                            {
                                unbound = string.Format("{0}/{1}", Ini.UnboundPdfFilePath, Path.GetFileName(pdffp));
                                if (File.Exists(unbound)) File.Delete(unbound);
                                File.Copy(pdffp, unbound);
                                File.Delete(pdffp);
                            }
                        }
                    }
                    else
                    {
                        if (Ini.DeleteDownloaded)
                        {
                            foreach (string pdffp in pdffilepathes)
                            {
                                if (File.Exists(pdffp)) File.Delete(pdffp);
                            }
                        }
                    }
                    Finish:;
                    //*/
                }
                //Повтроная обработка неприкрепленных pdf файлов
                if (Directory.Exists(Ini.UnboundPdfFilePath))
                {
                    string[] unPdffiles = Directory.GetFiles(Ini.UnboundPdfFilePath, "*.pdf");
                    foreach (string updf in unPdffiles)
                    {
                        string fNamePart = Path.GetFileName(updf).Split(new char[] { '_' })[0];
                        Console.WriteLine(updf);
                        AttachData aData = DBControl.GetAttachData(fNamePart);

                        if (attachPdf && aData != null)
                        {
                            string folder = string.Format("{0}-{1}/{2}", aData.ResDate.Year, Convert.ToString(aData.ResDate.Month).PadLeft(2, '0'), aData.PatientId);
                            string pdfDirectory = string.Format("{0}/{1}", Ini.PdfFilePath, folder);

                            if (!Directory.Exists(pdfDirectory))
                            {
                                DirectoryInfo dir = Directory.CreateDirectory(pdfDirectory);
                            }

                            string medialogfp = string.Format("{0}/{1}", pdfDirectory, Path.GetFileName(updf));
                            if (File.Exists(medialogfp)) File.Delete(medialogfp);
                            if (File.Exists(updf))
                            {
                                File.Copy(updf, medialogfp);
                                File.Delete(updf);
                                DBControl.AttachPdf(aData.MotconsuId, aData.PatientId, Path.GetFileName(updf), aData.ResDate, folder, Processing.Ini.RubricsId);
                            }

                        }

                    }
                }
                //EndOfWork: ;
            }

          
            //Загрузка справочников биматериалов, контейнеров, дополнительной информации
            //static public void GetInfo()

            //Вспомогательные функции

        }

        class Lorak
        {
            //Загрузка результатов
            static public void LoadResults()
            {

                DataTable dt = new DataTable();
                Encoding RusDOS = Encoding.GetEncoding(866);
                List<string> filelist = new List<string>();
                /*
                  dt = DbfFileFormat.Read(@"D:\data\8195553.DBF", "8195553");
                  foreach (DataRow row in dt.Rows)
                  {
                      var cells = row.ItemArray;
                      foreach (object cell in cells)
                          Console.Write("\t{0}", cell);
                      Console.WriteLine();
                  }

                  DbfFileFormat.Write(@"D:\data\test.DBF", dt, RusDOS);
                  //*/
                /*
                filelist = Files.FTPGetFileList(@"/%2ftmp/");
                foreach (var file in filelist)
                {
                    Console.WriteLine(file);

                    //Files.SaveStreamToFile("", Files.FTPDownload(@"/%2ftmp/" + file));
                    Stream fstream = Files.FTPDownload(@"/%2ftmp/" + file);
                    Files.SaveStreamToFile(@"D:/Data/" + file, fstream);
                }
                //*/
                //Stream dataStream =  Files.GetStreamFromFile(@"D:/Data/test41.dbf");

                //Files.SaveStreamToFile(@"D:/Data/" + "www.dbf", dataStream);

                //Files.FTPUpLoad(@"/%2ftmp/xxxx.dbf", dataStream);

                //Files.FTPDelete(@"/%2ftmp/" + "xx.dbf");

                /*
                using (WebClient client = new WebClient())
                {
                    client.Credentials = new NetworkCredential(Processing.mis, Processing.WPass);
                    //client.UploadFile(Processing.url + @"/%2ftmp/x.dbf", "STOR", @"D:/Data/test41.dbf");
                    //FileStream fileStream = File.OpenRead(@"D:/Data/test41.dbf");
                    Stream fileStream =  Files.GetStreamFromFile(@"D:/Data/test41.dbf");
                    byte[] buffer = new byte[fileStream.Length];
                    fileStream.Read(buffer, 0, buffer.Length);
                    client.UploadData(Processing.url + @"/%2ftmp/1234.dbf", "STOR", buffer);
                }
                //*/


            }

            //Выгрузка заказов
            static public void UploadOrders()
            {

            }
        }

        class KDL
        {
            //Загрузка результатов
            static public void LoadResults()
            {
                XDocument xdoc = null;
                int p_id = 0;

                bool attachPdf = Directory.Exists(Ini.PdfFilePath) && (Ini.RubricsId != 0);
                int motconsuId = 0;
                string resFolder = @"/%2fOutput";
                string pdfFolder = @"/%2fOutputPDF";
                List<string> filelist;
                //Stream fstream;
                //Stream pdfstream;
                if (Ini.TestMode == 2)
                {
                    Console.WriteLine("Test mode is on!");
                    filelist = Directory.GetFiles(Ini.ResultFilePath, "*.xml").ToList();
                }
                else
                {
                    try
                    {
                        filelist = Files.FTPGetFileList(resFolder, ".xml", true);
                        Console.WriteLine("На серввере {0} файлов для загрузки", filelist.Count);
                    }
                    catch (Exception ex)
                    {
                        filelist = null;
                        Files.WriteLogFile("Ошибка FTP: " + ex.ToString()); Console.WriteLine(ex.Message);
                        Environment.Exit(1);
                    }
                }
                
                //Загрузка и обработка файлов
                foreach (string fileName in filelist.Take(2))
                {
                    //Загрузка файла
                    Console.WriteLine("Загрузка результатов из файла {0}", fileName);
                    if (Ini.TestMode == 2)
                    {
                        xdoc = XDocument.Load(fileName);
                    }
                    else
                    {
                        try
                        {
                            //fstream = Files.FTPDownload(resFolder + @"/" + fileName);
                            //xdoc = XDocument.Load(fstream);
                            xdoc = Files.FTPDownloadXML(resFolder + @"/" + fileName, true);
                        }
                        catch (Exception ex)
                        {
                            Files.WriteLogFile("Ошибка FTP: " + ex.ToString()); Console.WriteLine(ex.Message);
                            Environment.Exit(1);
                        }
                    }

                    DataContext db = new DataContext(sqlConnectionStr);
                    try
                    {
                        XElement Patient = xdoc.Element("root").Element("Order").Element("Patient");
                        string[] FirstName = Patient.Element("FirstMiddleName").Value.Split(' ');

                        bool ok = true;

                        PatientName Name = new PatientName
                        {
                            LastName = Patient.Element("LastName").Value,
                            FirstName = FirstName[0],
                            MiddleName = FirstName.Count() > 1 ? FirstName[1] : string.Empty
                        };

                        DateTime PatBirthDate = Patient.Element("DOB") != null
                            ? DateTime.ParseExact(Patient.Element("DOB").Value, "yyyy-MM-dd", null)
                            : new DateTime(1990, 1, 1);

                        PatientBirthDate BirthDate = new PatientBirthDate
                        {
                            BirthYear = PatBirthDate.Year,
                            BirthMonth = PatBirthDate.Month,
                            BirthDay = PatBirthDate.Day
                        };

                        string KeyCode = xdoc.Element("root").Element("Order").Attribute("OrderID").Value;

                        //ok = DBControl.GetPatientID(db, KeyCode, out p_id);

                        ok = DBControl.GetPatientID(db, Name, BirthDate, fileName, out p_id);

                        XElement ResultsData = xdoc.Element("root").Element("Order").Element("Tests");

                        bool NewImport = DBControl.GetImpdataID(db, KeyCode, out int ImpdataId);

                        //Подсчет количества параметров исследований  Element("Analysis")
                        int resCount = (from xe in ResultsData.Elements("Test")
                                        select xe.Attribute("TestID")).Count();

                        Console.WriteLine("Количетсво строк {0}, id записи {1}", resCount, ImpdataId);

                        var resDatetime = string.Format("{0} {1}", ResultsData.Element("Test").Element("ValueDate").Value, ResultsData.Element("Test").Element("ValueTime").Value);
                        DateTime resdate = DateTime.ParseExact(resDatetime
                                        , "yyyy-MM-dd HH:mm:ss", null);

                        //Вставка в Impdata
                        if (NewImport)
                        {
                            ImpData impData = new ImpData()
                            {
                                ImpdataId = ImpdataId,
                                KEYCODE = KeyCode,
                                Nom = Name.LastName,
                                Prenom = Name.FirstName + " " + Name.MiddleName,
                                Date_Naissance = new DateTime(BirthDate.BirthYear, BirthDate.BirthMonth, BirthDate.BirthDay),
                                Date_Consultation = resdate,
                                Mesure = Mesure
                            };

                            //Console.WriteLine("Id\t {0}\t Key\t{1}\t Nom\t{2} \tPre|t{3}| \t Date\t{4} \tDD\t{5}\t MES\t{6}",
                            //impData.ImpdataId, impData.KEYCODE, impData.Nom, impData.Prenom, impData.Date_Naissance, impData.Date_Consultation, impData.Mesure);

                            DBControl.InsertImpdata(db, impData);
                        }

                        RestestData restestData;

                        //Комментарии к заявке
                        if (xdoc.Element("root").Element("Order").Element("OrderComment") != null)
                        {
                            var rComment = from xe in xdoc.Element("root").Element("Order").Elements("OrderComment")
                                           where xe.Value != string.Empty
                                           select xe.Value;

                            if (rComment.Count() > 0)
                            {
                                string ReqComment = string.Empty;
                                int i = 0;
                                foreach (var item in rComment)
                                {
                                    ReqComment += i == 0 ? item : Convert.ToString(Convert.ToChar(13)) + item;
                                    i++;
                                }

                                restestData = new RestestData
                                {
                                    M_VAL = ReqComment,
                                    Rescode = "ReqComment"
                                };

                                DBControl.InsertRestestItem(db, restestData, ImpdataId, resdate);
                            }
                        };

                        List<Pages> pagesList = (from xe in xdoc.Element("root").Element("Order").Element("Pages").Elements("Page").Elements("Section")
                                                 select new Pages
                                                 {
                                                     SectionId = xe.Attribute("SectionID").Value,
                                                     SectionName = xe.Element("SectionName").Value
                                                 }).ToList();

                        //Создание в Медиалоге групп исследований, если их еще нет
                        foreach (var p in pagesList)
                        {
                            int ResGroupNewId = DBControl.ResGroupCheckV2(db, p.SectionId, Ini.ResGroupId);
                            if (ResGroupNewId == 0)
                                DBControl.CreateResGroupV2(db, p.SectionName, p.SectionId, Ini.ResGroupId, out ResGroupNewId);
                        }

                        XElement testsX = xdoc.Element("root").Element("Order").Element("Tests");
                        //Построение лабораторных справочников и компоновка результата для импорта

                        List<AtachFile> attachFileList = new List<AtachFile>();

                        List<RestestData> results = new List<RestestData>();
                        foreach (XElement tstX in testsX.Nodes())
                        {
                            int testcount = (from xe in testsX.Elements("Test")
                                             select xe).Count();
                            if (testcount > 0)
                            {
                                string testCode = (tstX.Element("TestShortName") == null) ? tstX.Attribute("TestID").Value : tstX.Element("TestShortName").Value;
                                string testName = tstX.Element("TestName") != null ? tstX.Element("TestName").Value : string.Empty;
                                int methodId = DBControl.LabMethodCheck(db, testCode, out bool inStructure);
                                //Создание справочников
                                string unit = tstX.Element("Dimension") == null ? "" : tstX.Element("Dimension").Value;
                                string numGroup = tstX.Element("TestPosition").Element("SectionID").Value;
                                int resGroupId = DBControl.ResGroupCheckV2(db, numGroup, Ini.ResGroupId);
                                if (resGroupId != 0)
                                {
                                    if (methodId == 0)
                                        DBControl.BuildDitionary(db, testName, testCode
                                            , testCode, unit, resGroupId, out methodId);
                                    else
                                    {
                                        if (!inStructure)
                                            DBControl.BuildDitionary(db, methodId, testCode, unit, resGroupId);
                                    }
                                }
                                //Компановка результатов
                                results.Add(new RestestData
                                {
                                    VAL = tstX.Element("Value").Value,
                                    UNIT = unit,
                                    NORM_TEXT_REC = tstX.Element("NormalValue") != null ? tstX.Element("NormalValue").Value :
                                                               (tstX.Element("NormalValueMin") != null & tstX.Element("NormalValueMax") != null) ?
                                                               tstX.Element("NormalValueMin").Value + " - " + tstX.Element("NormalValueMax").Value : "",
                                    Comments = tstX.Element("TestComment") != null ? tstX.Element("TestComment").Value : "",
                                    MethodId = methodId
                                }
                                        );

                                //Обработка приложения в виде pdf
                                if (tstX.Element("Attach") != null)
                                {
                                    if (tstX.Element("Attach").Element("Body") != null)
                                        attachFileList.Add(new AtachFile
                                        {
                                            Name = tstX.Element("Attach").Attribute("FileName").Value,
                                            Body = tstX.Element("Attach").Element("Body").Value
                                        });
                                    //string attachFileName = tstX.Element("Attach").Attribute("FileName").Value;
                                    //string attachBody = tstX.Element("Attach").Element("Body").Value;
                                }
                            }
                        }

                        //foreach (var e in results)  Console.WriteLine("{0}\t{1}\t{2}\t{3}", e.MethodId, e.VAL, e.NORM_TEXT_REC, e.Comments);

                        //Вставки в DS_RESTESTS
                        DBControl.InsertRestestDataIDM(db, results, ImpdataId, resdate);

                        //Отправка результатов на тестового пациента
                        if (Ini.TestPatientID != 0) { ok = true; p_id = Ini.TestPatientID; }

                        string pdfFileName = fileName.Replace("xml", "PDF").Replace("XML", "PDF");

                        //Создание записи в ЭМК
                        if (ok)
                        {
                            DBControl.UpdateImpdata(db, ImpdataId, p_id);
                            //Запись в ЭМК
                            motconsuId = DBControl.InsertToEMCe(ImpdataId, p_id, resdate, NewImport);

                            //Дублирование комментариев в отдельную таблицу
                            DBControl.InsertToEMCTables(ImpdataId);

                            //Дублирование длинных комментариев к параметрам
                            if (Ini.CommentTable != null && Ini.CommentField != null) DBControl.ParamsComments(ImpdataId, motconsuId, p_id, Ini.CommentTable, Ini.CommentField);

                            if (Ini.TestMode == 0)
                            {
                                //Обработка pdf файла и прикрепление к ЭМК
                                try
                                {
                                    //pdfstream = Files.FTPDownload(pdfFolder + @"/" + pdfFileName);

                                    if (attachPdf && motconsuId != 0)
                                    {
                                        Console.WriteLine("Работа с pdf файлом");
                                        string folder = string.Format("{0}-{1}/{2}", resdate.Year, Convert.ToString(resdate.Month).PadLeft(2, '0'), p_id);
                                        string pdfDirectory = string.Format("{0}/{1}", Ini.PdfFilePath, folder);
                                        string medialogfp = string.Format("{0}/{1}", pdfDirectory, pdfFileName);
                                        //Console.WriteLine("pdfDirectory {0}", pdfDirectory);
                                        if (!Directory.Exists(pdfDirectory))
                                        {
                                            DirectoryInfo dir = Directory.CreateDirectory(pdfDirectory);
                                        }

                                        //Запись приложений из xml файла
                                        foreach (var att in attachFileList)
                                        {
                                            string attachFilePath = string.Format("{0}/{1}", pdfDirectory, att.Name);
                                            File.WriteAllBytes(attachFilePath, Convert.FromBase64String(att.Body));
                                            DBControl.AttachPdf(motconsuId, p_id, att.Name, resdate, folder, Processing.Ini.RubricsId);
                                        }

                                        //Копирование pdf на сервер Медиалога
                                        if (File.Exists(medialogfp)) File.Delete(medialogfp);
                                        //Console.WriteLine("pdfFolder {0}  pdfFileName {1}", pdfFolder, pdfFileName);
                                        if (Files.FTPFolderInfo(pdfFolder, true).ToLower().IndexOf(pdfFileName.ToLower()) != -1)
                                        {
                                            Console.WriteLine("pdf файл найден");
                                            //Files.SaveStreamToFile(medialogfp, pdfstream);
                                            Files.FTPDownloadFile(pdfFolder + @"/" + pdfFileName, medialogfp, true);
                                            DBControl.AttachPdf(motconsuId, p_id, pdfFileName, resdate, folder, Processing.Ini.RubricsId);
                                        }
                                        else
                                        {
                                            Console.WriteLine("Файл {1} не найден в {0}", pdfFolder, pdfFileName);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("Не возможно обработать pdf файл");
                                        if (Directory.Exists(Ini.ArchiveFilePath))
                                        {
                                            string pdfarchivefp = string.Format("{0}/{1}", Ini.ArchiveFilePath, pdfFileName);
                                            if (File.Exists(pdfarchivefp)) File.Delete(pdfarchivefp);
                                            //Files.SaveStreamToFile(pdfarchivefp, pdfstream);
                                            Files.FTPDownloadFile(pdfFolder + @"/" + pdfFileName, pdfarchivefp, true);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Files.WriteLogFile("Ошибка FTP: " + ex.ToString()); Console.WriteLine(ex.Message);
                                    Environment.Exit(1);
                                }
                            }
                        }

                        //Отправка файла в архив или удаление
                        if (Directory.Exists(Ini.ArchiveFilePath))
                        {
                            string archivefp = string.Format("{0}/{1}", Ini.ArchiveFilePath, fileName);
                            if (File.Exists(archivefp)) File.Delete(archivefp);
                            if (Ini.TestMode == 0)
                            {
                                Files.FTPDownloadFile(resFolder + @"/" + fileName, archivefp, true);
                                //Files.SaveStreamToFile(archivefp, fstream);
                            }
                        }

                        if (Ini.TestMode == 0)
                        {
                            //Удаление файлов с FTP-сервера
                            if (Ini.DeleteDownloaded)
                            {
                                if (Files.FTPFolderInfo(resFolder, true).ToLower().IndexOf(fileName.ToLower()) != -1) Files.FTPDelete(resFolder + @"/" + fileName, true);
                                if (Files.FTPFolderInfo(pdfFolder, true).ToLower().IndexOf(pdfFileName.ToLower()) != -1) Files.FTPDelete(pdfFolder + @"/" + pdfFileName, true);
                            }
                        }
                        
                    }
                    catch (Exception ex)
                    {
                        Files.WriteLogFile(string.Format("Ошибка разбора xml {0}: {1}", fileName, ex.ToString())); 
                        Console.WriteLine(ex.Message);
                        //Отправка файла в bad
                        string badFile = string.Format("{0}/{1}", Files.BadFilesDir, Path.GetFileName(fileName));
                        string pdfFileName = fileName.Replace("xml", "PDF").Replace("XML", "PDF");
                        if (Ini.TestMode ==2)
                        {
                            if (File.Exists(fileName))
                            {
                                File.Copy(fileName, badFile, true);
                                File.Delete(fileName);
                            }
                        }
                        if (Ini.TestMode == 0)
                        {
                            Files.FTPDownloadFile(resFolder + @"/" + fileName, badFile, true);
                            if (Ini.DeleteDownloaded)
                            {
                                Files.FTPDelete(resFolder + @"/" + fileName, true);
                                Files.FTPDelete(pdfFolder + @"/" + pdfFileName, true);
                            }
                        }
                    }
                    finally
                    {
                        db.Dispose();
                    }
                }
            }
            
            //Выгрузка заказов
            static public void UploadOrders()
            {
                CultureInfo culture = new CultureInfo("ru-RU"); //формат преобразования даты
                List<Request> RequestInfoList = new List<Request>();

                DBControl.GetOrdersList(0, 1, out RequestInfoList);

                if (RequestInfoList.Count == 0) Console.WriteLine("Нет заказов для отправки");
                else Console.WriteLine("Готово к отправке {0} заказов", RequestInfoList.Count);
                //Формирование файлов заказов
                if (RequestInfoList.Count != 0)
                {
                    Console.WriteLine("Создание и выгрузка готовых заказов...");
                    //*
                    foreach (Request qRequest in RequestInfoList)
                    {
                        //Получение данных заказа из базы данных
                        DBControl.GetOrderInfoV4(qRequest.ID, out PatientInfo qPatient, out List<AssayInfo> qAssay, out List<Order> qOrder);
                        //Группировка по услугам в заказе
                        var gOrder = from order in qOrder
                                     group order by new
                                     {
                                         order.Name,
                                         order.Code,
                                         order.BioType,
                                         order.GroupOption
                                     } into groupOrder
                                     select new
                                     {
                                         Name = groupOrder.Key.Name,
                                         Code = groupOrder.Key.Code,
                                         BioType = groupOrder.Key.BioType,
                                         GroupOption = groupOrder.Key.GroupOption
                                     };
                        //Получение доп. информации
                        DBControl.GetAddInfo(qRequest.ID, out List<AddInfo> AddInfoList);

                        //Создание тела XML сообщения для создания заказа в ЛИС 
                        XElement messagebody = new XElement("root"
                            , new XElement("Header"
                                , new XElement("Version", "4")
                                , new XElement("FileDate", qRequest.Date.ToString("yyyy-MM-dd"))
                                , new XElement("FileTime", qRequest.Date.ToString("HH:mm:ss"))
                                , new XElement("GenTime", "0")
                                , new XElement("FileType", "Order")
                                , new XElement("LaboratoryID", "kdltest_united")
                                , new XElement("LaboratoryName", "КДЛ-ТЕСТ")
                                , new XElement("ClinicID", Ini.HospitalID)
                                , new XElement("ClinicName", Ini.HospitalName)
                                , new XElement("Direction", "FromClinic")
                                            )
                            , new XElement("Order", new XAttribute("OrderID", qRequest.Nr)
                                    , new XElement("OrderComment", qRequest.Comment)
                                    , new XElement("Patient",
                                             new XElement("LastName", qPatient.LastName),
                                             new XElement("FirstMiddleName", string.Format("{0} {1}", qPatient.FirstName, qPatient.MiddleName)),
                                             new XElement("DOB", qPatient.BirthDate.ToString("yyyy-MM-dd")),
                                             new XElement("Sex", qPatient.Gender == 0 ? "M" : qPatient.Gender == 1 ? "F" : "U")
                                                    )
                                    , new XElement("Tests",
                                              from test in gOrder
                                              select new XElement("Test", new XElement("TestShortName", test.Code),
                                                            new XElement("TestName", test.Name),
                                                            from sample in qAssay
                                                            where sample.ProfileCode == test.Code
                                                                  && test.BioType != 0
                                                            select new XElement("locus", sample.BiomaterialCode),   // test.Obligatory == 0 ? sample.BiomaterialCode : string.Empty
                                                                    from info in AddInfoList
                                                                    where info.ProfileCode == test.Code 
                                                                        && info.GroupOption == test.GroupOption
                                                                    select new XElement("DopInfo",
                                                                    new XElement("Info", new XAttribute("ID", info.InfoId)
                                                                            , new XElement("Value", info.Value)
                                                                                )      //Info
                                                                        ) //DopInfo
                                                                   ) //Test
                                                    ) //Tests
                                            )//Order
                                ); //root

                        XDocument xdoc = new XDocument(new XDeclaration("1.0", "windows-1251", ""), messagebody);
                        
                        //Запись файла заказов
                        if (Directory.Exists(Ini.OrderFilePath))
                        {
                            string filename = string.Format(@"{1}\{0}.xml", qRequest.Nr, Ini.OrderFilePath);
                            xdoc.Save(filename);
                        }

                        if (Ini.TestMode == 0)
                        {
                            //Загрузка файла на FTP-сервер
                            Stream xdocStream = new MemoryStream();
                            xdoc.Save(xdocStream);
                            xdocStream.Position = 0;
                            Files.FTPUpLoad(@"/%2fInput/" + qRequest.Nr + ".xml", xdocStream);

                            Console.WriteLine("Файл {0}.xml выгружен на сервер", qRequest.Nr);
                            //Изменение статуса заказа
                            DBControl.UpdateOrderInfo(qRequest.ID, 2, qRequest.Nr.ToString());
                        }
                    }
                    //*/
                    /*
                    //Проверка статуса заказов
                    Console.WriteLine("Проверка статуса заказов...");
                    Thread.Sleep(10000);

                    string folderInfo = Files.FTPFolderInfo(@"/%2fLogs/");
                    Stream fstream;
                    XDocument report;
                    foreach (Request qRequest in RequestInfoList)
                    {
                        //Загрузка файла отчета
                        if (folderInfo.IndexOf(qRequest.Code, StringComparison.OrdinalIgnoreCase) != -1)
                        {
                            fstream = Files.FTPDownload(@"/%2fLogs/" + @"/" + qRequest.Code + ".xml");
                            report = XDocument.Load(fstream);
                            List<Report> qReport = (from xe in report.Element("Order").Elements("TestsOK").Elements("TestsOKItem")
                                                    where xe.Value != string.Empty
                                                    select new Report
                                                    {
                                                        OrderId = xe.Parent.Parent.Attribute("OrderID").Value,
                                                        Message = xe.Value
                                                    }).ToList();
                            foreach (var r in qReport)
                            {
                                Console.WriteLine("принят тест {1} в заказе {0}", r.OrderId, r.Message);
                            }

                            qReport = (from xe in report.Element("Order").Elements("Errors").Elements("ErrorsItem")
                                       where xe.Value != string.Empty
                                       select new Report
                                       {
                                           OrderId = xe.Parent.Parent.Attribute("OrderID").Value,
                                           Message = xe.Value
                                       }).ToList();
                            foreach (var r in qReport)
                            {
                                Console.WriteLine("Ошибка: \"{1}\" в заказе {0}", r.OrderId, r.Message);
                                Files.WriteLogFile(string.Format("Ошибка: \"{1}\" в заказе {0}", r.OrderId, r.Message));
                            }
                            //Изменение статуса заказа
                            DBControl.UpdateOrderState(qRequest.ID, 1);
                            //Удаление файла логов
                           // Files.FTPDelete(@"/%2fLogs/" + @"/" + qRequest.Code + ".xml");
                        }
                    }
                    Console.WriteLine("Статус заказов проверен");
                    //*/
                }
                //*/
            }
        }

        class Ariadna
        {
            //Загрузка результатов
            static public void LoadResults()
            {
                XDocument xdoc;
                int p_id = 0;

                string[] resfiles = Directory.GetFiles(Ini.ResultFilePath, "*.xml");
                int motconsuId = 0;

                foreach (string filepath in resfiles)
                {
                    DataContext db = new DataContext(sqlConnectionStr);
                    Table<Impdata> TImpdata = db.GetTable<Impdata>();

                    xdoc = XDocument.Load(filepath);
                    Console.WriteLine("Загрузка результатов из файла {0}", filepath);

                    XElement xobservation = xdoc.Element("MedapLisObservationReport").Element("Observation");

  
                    XElement xpatient = xobservation.Element("Patient");
                    string StrPatID = xpatient.Element("ExternalID") != null ? xpatient.Element("ExternalID").Value : "";
                    string KeyCode = xobservation.Element("LisOrderID").Value;
                    int HisOrderId = Convert.ToInt32( xobservation.Element("HisOrderID") != null ? xobservation.Element("HisOrderID").Value : "0");
                    

                    bool ok = true;

                    PatientName Name = new PatientName();

                    //Дата рождения пациента из файла
                    DateTime PatBirthDate = xpatient.Element("BirthDate") != null ? DateTime.ParseExact(xpatient.Element("BirthDate").Value, "yyyy-MM-dd", null) : new DateTime(1900,1,1);

                    PatientBirthDate BirthDate = new PatientBirthDate
                    {
                        BirthYear = PatBirthDate.Year,
                        BirthMonth = PatBirthDate.Month,
                        BirthDay = PatBirthDate.Day
                    };

                    //Идентификация пациента по данным из файла
                    if (!DBControl.GetPatientID(db, StrPatID, filepath, HisOrderId, out p_id))
                    {
                        ok = false;
                        Name = new PatientName
                        {
                            LastName = xpatient.Element("FamilyName") != null ? xpatient.Element("FamilyName").Value : "",
                            FirstName = xpatient.Element("GivenName") != null ? xpatient.Element("GivenName").Value : "",
                            MiddleName = xpatient.Element("MiddleName") != null ? xpatient.Element("MiddleName").Value : ""
                        };

                        ok = DBControl.GetPatientID(db, Name, BirthDate, filepath, out p_id);
                    }

                    bool NewImport = DBControl.GetImpdataID(db, KeyCode, out int ImpdataId);


                    if (xobservation.Element("ObservationReport") != null)
                    {
                        //Подсчет количесва параметров исследований
                        int resCount = (from xe in xobservation.Elements("ObservationReport").Elements("ReportGroup").
                                       Elements("Results").Elements("ObservationResult")
                                        where xe.Element("ResCode") != null
                                        select xe.Element("ResCode")).Count();

                        DateTime resdate = Convert.ToDateTime(xobservation.Element("FinishDate").Value);

                        if (NewImport)
                        {
                            Console.WriteLine("Количетсво строк {0}, id записи {1}", resCount, ImpdataId);
                            var insetImpdata = from xe in xobservation.Elements("OrderInfo")
                                               select new Impdata
                                               {
                                                   ImpdataId = ImpdataId,
                                                   KEYCODE = KeyCode,
                                                   Nom = Name.LastName,
                                                   Prenom = Name.FirstName + " " + Name.MiddleName,
                                                   Date_Naissance = PatBirthDate,
                                                   Date_Consultation = resdate,
                                                   Mesure = Mesure
                                               };
                            //Вставка в Impdata
                            TImpdata.InsertAllOnSubmit(insetImpdata);
                            db.SubmitChanges();
                        }
                        //Выборка для вставки в DS_RESTESTS
 
                        List<RestestData> qRestests = (from xe in xobservation.Element("ObservationReport").Element("ReportGroup").
                                                        Elements("Results").Elements("ObservationResult")
                                                       where xe.Element("ResCode") != null
                                                       select new RestestData
                                                       {
                                                           VAL = xe.Element("ResultValue") == null ? "" : xe.Element("ResultValue").Value,
                                                           UNIT = xe.Element("Unit") == null ? "" : xe.Element("Unit").Value,
                                                           Rescode = xe.Element("ResCode").Value,    //ResCode ? MeasurementCode ?
                                                           ResName = xe.Element("MeasurementName").Value,
                                                           NORM_TEXT_REC = xe.Element("NormText") == null ? "" : xe.Element("NormText").Value,
                                                           State = xe.Element("Pathology") == null ? 2 : (xe.Element("Pathology").Value == "0" ? 0:2)
                                                       }).ToList();

                        DBControl.InsertRestestData(db, qRestests, ImpdataId, resdate);

                        //Отправка результатов на тестового пациента
                        if (Ini.TestPatientID != 0) { ok = true; p_id = Ini.TestPatientID; }

                        //Создание записи в ЭМК
                        if (ok)
                        {
                            DBControl.UpdateImpdata(db, ImpdataId, p_id);
                            
                            //Запись в ЭМК
                            motconsuId = DBControl.InsertToEMCe(ImpdataId, p_id, resdate, NewImport);
                            
                            //Дублирование комментариев в отдельную таблицу
                            DBControl.InsertToEMCTables(ImpdataId);

                            //Дублирование длинных комментариев к параметрам
                            if (Ini.CommentTable != null && Ini.CommentField != null)
                                DBControl.ParamsComments(ImpdataId, motconsuId, p_id, Ini.CommentTable, Ini.CommentField);
                        }
                        db.Dispose();
                    }

                    //Копирование файла в архив
                    if (Directory.Exists(Ini.ArchiveFilePath))
                    {
                        string fp = string.Format("{0}/{1}", Ini.ArchiveFilePath, Path.GetFileName(filepath));
                        if (File.Exists(fp)) File.Delete(fp);
                        File.Copy(filepath, fp);
                    }

                    //Удаление файла из текущей директории
                    if (Ini.DeleteDownloaded)
                    {
                        File.Delete(filepath);
                        string filepathmd5 = filepath.Replace("xml", "md5");
                        File.Delete(filepathmd5);
                    }
                }
            }

            //Выгрузка заказов
            static public void UploadOrders()
            {
                CultureInfo culture = new CultureInfo("ru-RU"); //формат преобразования даты
                List<Request> RequestInfoList = new List<Request>();

                DBControl.GetOrdersList(0, 5, out RequestInfoList);

                if (RequestInfoList.Count == 0) Console.WriteLine("Нет направлений для заказов");
                else Console.WriteLine("Создано {0} заказов", RequestInfoList.Count);

                ///*
                if (RequestInfoList.Count != 0)
                {
                    Console.WriteLine("Отправка готовых заказов...");
                    XElement messagebody = null;
                    foreach (Request qRequest in RequestInfoList)
                    {
                        //Получение данных заказа из базы данных
                        DBControl.GetOrderInfoV4(qRequest.ID, out PatientInfo qPatient, out List<AssayInfo> qAssay, out List<Order> qOrder);

                        //Создание тела XML сообщения для создания заказа в ЛИС 
                        messagebody = new XElement("MedapLisObservationRequest",
                                from assay in qAssay
                                select new XElement("Observation"
                                    , new XElement("ID", qRequest.Nr)
                                    , new XElement("IDs", assay.Barcode)
                                    , new XElement("RegDate", qRequest.Date.ToString("yyyy-MM-ddTHH:mm:ss"))
                                    , new XElement("HisOrderID", qRequest.ID)
                                    , new XElement("HisSampleID", assay.BarcodeId)
                                    , new XElement("OrderID", qRequest.ID)
                                    , new XElement("OrderDate", qRequest.Date.ToString("yyyy-MM-dd"))
                                    , new XElement("HisSampleID", assay.BarcodeId)
                                    , new XElement("SpecimenTypeID", assay.BiomaterialCode)
                                    , new XElement("SpecimenType", assay.BiomaterialLabel)
                                    , new XElement("CollectDate", qRequest.Date.ToString("yyyy-MM-ddTHH:mm:ss"))
                                   , new XElement("Patient",
                                                new XElement("RegCode", qPatient.ID),
                                                new XElement("FamilyName", qPatient.LastName),
                                                new XElement("GivenName", qPatient.FirstName),
                                                new XElement("MiddleName", qPatient.MiddleName),
                                                new XElement("BirthDate", qPatient.BirthDate.ToString("yyyy-MM-dd")),
                                                new XElement("Gender", qPatient.Gender == 0 ? "M" : qPatient.Gender == 1 ? "F" : " "),
                                                new XElement("ExternalID", qPatient.ID)
                                                       )  //Patient
                                   , new XElement("OrderingInstitution",
                                                new XElement("ID", qRequest.DepId),
                                                new XElement("FullName", qRequest.DepLabel),
                                                new XElement("Code", qRequest.DepCode),
                                                new XElement("Physician",
                                                       new XElement("GivenName", "")
                                                             )
                                                  )
                                   , new XElement("OrderInfo",
                                            from test in qOrder
                                            where test.BarcodeId == assay.BarcodeId
                                            select new XElement("item",
                                                  new XElement("ServiceCode", test.ExternalId),
                                                  new XElement("ServiceID", test.Code),
                                                  new XElement("Service", test.Name)
                                                            )   //item
                                                    )   //OrderInfo
                                            )  // Observation
                                   ); //MedapLisObservationRequest

                        string filename = (DateTime.Now.Year - 2000).ToString().PadLeft(2,'0') + DateTime.Now.DayOfYear.ToString().PadLeft(3, '0')
                                +  Message.DecToZex(DBControl.GetCounterValue("msg_num", 1), 36).PadLeft(3,'0');
                        string md5Content = string.Format("MD5 ({0}.XML) = {1}", filename, Files.GetMD5(filename + ".XML"));
                        Files.WriteToFile(string.Format("{0}/{1}.md5", Ini.OrderFilePath, filename), md5Content, false);
                        XDocument xdoc = new XDocument(new XDeclaration("1.0", "utf-8", ""), messagebody);
                        xdoc.Save(string.Format("{0}/{1}.xml", Ini.OrderFilePath, filename));

                        //Изменение статуса заказа
                        DBControl.UpdateOrderState(qRequest.ID, 1);
                    }
                }
                //*/

            }

        }

        class Helix
        {
            static public void LoadResults()
            {
                XDocument xdoc = null;
                int p_id = 0;

                string[] resfiles = Directory.GetFiles(Ini.ResultFilePath, "*.xml");
                bool attachPdf = Directory.Exists(Ini.PdfFilePath) && (Ini.RubricsId != 0);
                int motconsuId = 0;
                DateTime resdate = DateTime.Now;
                foreach (string filepath in resfiles)
                {
                    DataContext db = new DataContext(sqlConnectionStr);
                    Table<Impdata> TImpdata = db.GetTable<Impdata>();

                    xdoc = XDocument.Load(filepath);
                    Console.WriteLine("Загрузка результатов из файла {0}", filepath);

                    XElement orderX = xdoc.Element("response").Element("Order");

                    bool ok = true;

                    PatientName Name = new PatientName
                    {
                        LastName = orderX.Attribute("last_name").Value,
                        FirstName = orderX.Attribute("first_name").Value,
                        MiddleName = orderX.Attribute("middle_name").Value
                    };

                    DateTime PatBirthDate = DateTime.Parse(orderX.Attribute("birthdate").Value);

                    PatientBirthDate BirthDate = new PatientBirthDate
                    {
                        BirthYear = PatBirthDate.Year,
                        BirthMonth = PatBirthDate.Month,
                        BirthDay = PatBirthDate.Day
                    };

                    ok = DBControl.GetPatientID(db, Name, BirthDate, filepath, out p_id);


                    XElement sampleX = orderX.Element("Sample");

                    string KeyCode = orderX.Attribute("lis_order_id").Value;
                    string misSampleId = sampleX.Attribute("mis_sample_id").Value;
                    string pdffp = string.Format(@"{0}\{1}.pdf", Ini.ResultFilePath, misSampleId);

                    string status = sampleX.Attribute("status").Value;
                    
                    if (status == "A")
                    {

                        bool NewImport = DBControl.GetImpdataID(db, KeyCode, out int ImpdataId);

                        //Подсчет количества исследований  Element("Test")
                        int resCount = 0;
                        if (sampleX.Elements("Test") != null)
                            resCount = (from xe in sampleX.Elements("Test")
                                        select xe.Attribute("Reported_name")).Count();

                        Console.WriteLine("Количетсво строк {0}, id импорта {1}", resCount, ImpdataId);

                        resdate = DateTime.Parse(sampleX.Attribute("authorized_on").Value);

                        //Вставка в Impdata
                        if (NewImport)
                        {
                            ImpData impData = new ImpData()
                            {
                                ImpdataId = ImpdataId,
                                KEYCODE = KeyCode,
                                Nom = Name.LastName,
                                Prenom = Name.FirstName + " " + Name.MiddleName,
                                Date_Naissance = new DateTime(BirthDate.BirthYear, BirthDate.BirthMonth, BirthDate.BirthDay),
                                Date_Consultation = resdate,
                                Mesure = Mesure,
                                LabCode = misSampleId
                            };
                            DBControl.InsertImpdata(db, impData);
                        }



                        List<RestestData> results = new List<RestestData>();
                        //Комментарии к заявке
                        if (sampleX.Attribute("comment") != null)
                        {
                            int methodId = DBControl.LabMethodCheck(db, "ReqComment");
                            //Создание справочников
                            if (methodId == 0)
                                DBControl.BuildDitionary(db, "Комментарии к заявке", "ReqComment", "", "", Ini.ResGroupId, out methodId);
                            string val = sampleX.Attribute("comment").Value;
                            if (val != string.Empty)
                            {
                                //Компановка результатов
                                results.Add(new RestestData
                                {
                                    VAL = val,
                                    MethodId = methodId
                                }
                                        );
                            }
                        }

                        //Биоматериал
                        if (sampleX.Attribute("Biomaterial") != null) 
                        {
                            int methodId = DBControl.LabMethodCheck(db, "Biomaterial");
                            //Создание справочников
                            if (methodId == 0)
                                DBControl.BuildDitionary(db, "Биоматериал", "Biomaterial", "", "", Ini.ResGroupId, out methodId);
                            string val = sampleX.Attribute("Biomaterial").Value;
                            if (val != string.Empty)
                            {
                                //Компановка результатов
                                results.Add(new RestestData
                                {
                                    VAL = val,
                                    MethodId = methodId
                                }
                                        );
                            }
                        }

                        //Построение лабораторных справочников и компоновка результата для импорта
                        foreach (XElement testX in sampleX.Nodes())
                        {
                            int testcount = (from xe in sampleX.Elements("Test")
                                             select xe).Count();
                            if (testcount > 0) //Обычные тесты, не микробиология
                            {
                                resCount = (from xe in testX.Elements("Result")
                                            select xe.Attribute("Reported_name")).Count();
                                if (resCount == 1)   //Простой тест
                                {
                                    int methodId = DBControl.LabMethodCheck(db, testX.Attribute("Code").Value);
                                    //Создание справочников
                                    string unit = testX.Element("Result").Attribute("UnitName") != null ? testX.Element("Result").Attribute("UnitName").Value : "";
                                    if (methodId == 0)
                                        DBControl.BuildDitionary(db, testX.Attribute("Reported_name").Value, testX.Attribute("Code").Value
                                          , testX.Attribute("Item_code").Value, unit, Ini.ResGroupId, out methodId);
                                    //Компановка результатов
                                    string ref_lo = testX.Element("Result").Attribute("ref_lo") != null ? testX.Element("Result").Attribute("ref_lo").Value : string.Empty;
                                    string ref_hi = testX.Element("Result").Attribute("ref_hi") != null ? testX.Element("Result").Attribute("ref_hi").Value : string.Empty;
                                    results.Add(new RestestData
                                    {
                                        VAL = testX.Element("Result").Attribute("Value").Value,
                                        UNIT = unit,
                                        NORM_TEXT_REC = ref_lo != string.Empty ? ref_lo + " - " + ref_hi : ref_hi,
                                        Comments = testX.Attribute("Comment") != null ? testX.Attribute("Comment").Value : null,
                                        MethodId = methodId
                                    }
                                            );
                                }
                                if (resCount > 1)    //Составной тест
                                {
                                    int ResGroupNewId = DBControl.ResGroupCheck(db, testX.Attribute("Item_code").Value, Ini.ResGroupId);
                                    //Создание новой группы исследований
                                    if (ResGroupNewId == 0)
                                        DBControl.CreateResGroup(db, testX.Attribute("Reported_name").Value, testX.Attribute("Item_code").Value, Ini.ResGroupId, out ResGroupNewId);
                                    int k = 0;
                                    foreach (XElement resultX in testX.Nodes())
                                    {
                                        string ncode = testX.Attribute("Item_code").Value + " " + resultX.Attribute("Name").Value;
                                        int methodId = DBControl.LabMethodCheck(db, ncode);
                                        //Создание справочников
                                        string unit = resultX.Attribute("UnitName") != null ? resultX.Attribute("UnitName").Value : "";
                                        k++;
                                        if (methodId == 0)
                                            DBControl.BuildDitionary(db, resultX.Attribute("Name").Value, ncode, string.Format("{0}-{1}", testX.Attribute("Item_code").Value, k)
                                            , unit, ResGroupNewId, out methodId);
                                        //Компановка результатов
                                        string ref_lo = resultX.Attribute("ref_lo") != null ? resultX.Attribute("ref_lo").Value : string.Empty;
                                        string ref_hi = resultX.Attribute("ref_hi") != null ? resultX.Attribute("ref_hi").Value : string.Empty;
                                        results.Add(new RestestData
                                        {
                                            VAL = resultX.Attribute("Value").Value,
                                            UNIT = unit,
                                            NORM_TEXT_REC = ref_lo != string.Empty ? ref_lo + " - " + ref_hi : ref_hi,
                                            MethodId = methodId
                                        }
                                            );
                                    }
                                }

                            }

                            int testmiccount = (from xe in sampleX.Elements("TestMicroorganism")
                                                select xe).Count();
                            if (testmiccount > 0)   //Микробиология и чувствительность к антибиотикам
                            {
                                //Комментарии к исследованию
                                if (sampleX.Element("TestMicroorganism").Attribute("Comment") != null)
                                {
                                    int methodId = DBControl.LabMethodCheck(db, "AnaComment");
                                    //Создание справочников
                                    if (methodId == 0)
                                        DBControl.BuildDitionary(db, "Комментарии к исследованию", "AnaComment", "", "", Ini.CultGroupId, out methodId);

                                    string val = sampleX.Element("TestMicroorganism").Attribute("Comment").Value;
                                    if (val != string.Empty)
                                    {
                                        //Компановка результатов
                                        results.Add(new RestestData
                                        {
                                            VAL = "-",
                                            Comments = val,
                                            MethodId = methodId
                                        }
                                                );
                                    }
                                }

                                //Микроорганизмы
                                var microorganism = from xe in testX.Elements()
                                                    where xe.Name == "Result"
                                                    select xe;

                                foreach (var mic in microorganism)
                                {
                                    string countvalue = mic.Attribute("CountValue").Value;
                                    if (mic.Attribute("Name").Value == "Выделенная флора" || countvalue != string.Empty)
                                    {

                                        int methodId = DBControl.LabMethodCheck(db, mic.Attribute("Value").Value);
                                        //Создание справочников
                                        if (methodId == 0)
                                            DBControl.BuildDitionary(db, mic.Attribute("Value").Value, mic.Attribute("Value").Value
                                              , string.Empty, string.Empty, Ini.CultGroupId, out methodId);
                                        //Компановка результатов
                                        results.Add(new RestestData
                                        {
                                            VAL = countvalue,
                                            NORM_TEXT_REC = mic.Attribute("NormValue") != null ? mic.Attribute("NormValue").Value : string.Empty,
                                            Comments = mic.Attribute("Pathogenicity") != null ? mic.Attribute("Pathogenicity").Value : string.Empty,
                                            MethodId = methodId
                                        }
                                                );
                                    }
                                    else
                                    {
                                        string ncode = testX.Attribute("Item_code").Value + " " + mic.Attribute("Name").Value;
                                        int methodId = DBControl.LabMethodCheck(db, ncode);
                                        //Создание справочников
                                        if (methodId == 0)
                                            DBControl.BuildDitionary(db, mic.Attribute("Name").Value, ncode, testX.Attribute("Item_code").Value
                                            , string.Empty, Ini.CultGroupId, out methodId);
                                        string val = mic.Attribute("Value").Value;
                                        val = val != string.Empty ? val : mic.Attribute("CountValue").Value;
                                        if (val != string.Empty)
                                        {
                                            //Компановка результатов
                                            results.Add(new RestestData
                                            {
                                                VAL = val,
                                                MethodId = methodId
                                            }
                                                );
                                        }
                                    }
                                }
                                //Антибиотики
                                var antibioticRes = from xe in testX.Elements()
                                                    where xe.Name == "Antibiotic"
                                                    select xe;
                                int item = 0;
                                foreach (var antRes in antibioticRes)
                                {
                                    item++;
                                    var antibioticName = from xe in antRes.Elements()
                                                         where xe.Name == "ResultAntibiotic"
                                                         select xe;
                                    foreach (var antName in antibioticName)
                                    {
                                        string code = antName.Attribute("Name").Value == "Вид м/о" ? "АБ " + "Вид м/о" : antName.Attribute("Name").Value;
                                        int methodId = DBControl.LabMethodCheck(db, code);
                                        //Создание справочников
                                        string unit = antName.Attribute("UnitName") != null ? antName.Attribute("UnitName").Value : "";
                                        if (methodId == 0)
                                            DBControl.BuildDitionary(db, antName.Attribute("Name").Value, code
                                              , antRes.Attribute("Item_code").Value, unit, Ini.AntbGroupId, out methodId);
                                        //Компановка результатов
                                        string val = antName.Attribute("Sensitivity") != null
                                                    ? string.Format("{0} | {1} {2}", antName.Attribute("Sensitivity").Value, antName.Attribute("Value").Value, unit)
                                                    : antName.Attribute("Value").Value;
                                        results.Add(new RestestData
                                        {
                                            VAL = val,
                                            UNIT = unit,
                                            Item = item,
                                            MethodId = methodId
                                        }
                                                );
                                    }
                                }
                                //Бактериофаги
                                var bacteriophageRes = from xe in testX.Elements()
                                                       where xe.Name == "Bacteriophage"
                                                       select xe;
                                item = 0;
                                foreach (var bacRes in bacteriophageRes)
                                {
                                    item++;
                                    var bacteriophageName = from xe in bacRes.Elements()
                                                            where xe.Name == "ResultBacteriophage"
                                                            select xe;
                                    foreach (var bacName in bacteriophageName)
                                    {
                                        string code = bacName.Attribute("Name").Value == "Вид м/о" ? "ФАГИ " + "Вид м/о" : bacName.Attribute("Name").Value;
                                        int methodId = DBControl.LabMethodCheck(db, code);
                                        //Создание справочников
                                        string unit = bacName.Attribute("UnitName") != null ? bacName.Attribute("UnitName").Value : "";
                                        if (methodId == 0)
                                            DBControl.BuildDitionary(db, bacName.Attribute("Name").Value, code
                                              , bacRes.Attribute("Item_code").Value, unit, Ini.BctrGroupId, out methodId);
                                        //Компановка результатов
                                        string val = bacName.Attribute("Sensitivity") != null
                                                      ? string.Format("{0} | {1} {2}", bacName.Attribute("Sensitivity").Value, bacName.Attribute("Value").Value, unit)
                                                      : bacName.Attribute("Value").Value;
                                        results.Add(new RestestData
                                        {
                                            VAL = val,
                                            UNIT = unit,
                                            Item = item,
                                            MethodId = methodId
                                        }
                                                );
                                    }
                                }

                            }
                        }

                        //foreach (var e in results)  Console.WriteLine("{0}\t{1}\t{2}\t{3}", e.MethodId, e.VAL, e.NORM_TEXT_REC, e.Comments);
                        //Отправка результатов на тестового пациента
                        if (Ini.TestPatientID != 0) { ok = true; p_id = Ini.TestPatientID; }

                        //Вставки в DS_RESTESTS
                        DBControl.InsertRestestDataIDM(db, results, ImpdataId, resdate);
                        if (ok)
                        {
                            DBControl.UpdateImpdata(db, ImpdataId, p_id);
                            //Запись в ЭМК
                            motconsuId = DBControl.InsertToEMCe(ImpdataId, p_id, resdate, NewImport);

                            //Дублирование комментариев в отдельную таблицу
                            DBControl.InsertToEMCTables(ImpdataId);

                            //Дублирование длинных комментариев к параметрам
                            if (Ini.CommentTable != null && Ini.CommentField != null)
                                DBControl.ParamsComments(ImpdataId, motconsuId, p_id, Ini.CommentTable, Ini.CommentField);

                            //Обработка pdf файла и прикрепление к ЭМК
                            if (attachPdf && motconsuId != 0)
                            {
                                string folder = string.Format("{0}-{1}/{2}", resdate.Year, Convert.ToString(resdate.Month).PadLeft(2, '0'), p_id);
                                string pdfDirectory = string.Format("{0}/{1}", Ini.PdfFilePath, folder);

                                if (!Directory.Exists(pdfDirectory))
                                {
                                    DirectoryInfo dir = Directory.CreateDirectory(pdfDirectory);
                                }


                                if (File.Exists(pdffp))
                                {
                                    string medialogfp = string.Format("{0}/{1}", pdfDirectory, Path.GetFileName(pdffp));
                                    if (File.Exists(medialogfp)) File.Delete(medialogfp);
                                    if (File.Exists(pdffp))
                                    {
                                        File.Copy(pdffp, medialogfp);
                                        File.Delete(pdffp);
                                        DBControl.AttachPdf(motconsuId, p_id, Path.GetFileName(pdffp), resdate, folder, Processing.Ini.RubricsId);
                                    }
                                }
                            }

                        }
                        db.Dispose();

                        //Отправка файла во временную директорию или удаление
                        if (Directory.Exists(Ini.ArchiveFilePath))
                        {
                            string archivefp = string.Format("{0}/{1}", Ini.ArchiveFilePath, Path.GetFileName(filepath));
                            if (File.Exists(archivefp)) File.Delete(archivefp);
                            File.Copy(filepath, archivefp);
                            File.Delete(filepath);
                        }
                        else
                        {
                            if (Ini.DeleteDownloaded)
                            {
                                File.Delete(filepath);
                            }
                        }

                        if (Directory.Exists(Ini.UnboundPdfFilePath))
                        {
                            string unbound = string.Format("{0}/{1}", Ini.UnboundPdfFilePath, Path.GetFileName(filepath));
                            if (File.Exists(pdffp))
                            {
                                unbound = string.Format("{0}/{1}", Ini.UnboundPdfFilePath, Path.GetFileName(pdffp));
                                if (File.Exists(unbound)) File.Delete(unbound);
                                File.Copy(pdffp, unbound);
                                File.Delete(pdffp);
                            }

                        }
                        else
                        {
                            if (File.Exists(pdffp)) File.Delete(pdffp);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Статус образца: Отменен");
                    }
                }
                //Повтроная обработка неприкрепленных pdf файлов
                if (Directory.Exists(Ini.UnboundPdfFilePath))
                {
                    string[] unPdffiles = Directory.GetFiles(Ini.UnboundPdfFilePath, "*.pdf");
                    foreach (string updf in unPdffiles)
                    {
                        string fNamePart = Path.GetFileName(updf).Split(new char[] { '.' })[0];
                        Console.WriteLine(updf);
                        AttachData aData = DBControl.GetAttachDataS(fNamePart);

                        if (attachPdf && aData != null)
                        {
                            string folder = string.Format("{0}-{1}/{2}", aData.ResDate.Year, Convert.ToString(aData.ResDate.Month).PadLeft(2, '0'), aData.PatientId);
                            string pdfDirectory = string.Format("{0}/{1}", Ini.PdfFilePath, folder);

                            if (!Directory.Exists(pdfDirectory))
                            {
                                DirectoryInfo dir = Directory.CreateDirectory(pdfDirectory);
                            }

                            string medialogfp = string.Format("{0}/{1}", pdfDirectory, Path.GetFileName(updf));
                            if (File.Exists(medialogfp)) File.Delete(medialogfp);
                            if (File.Exists(updf))
                            {
                                File.Copy(updf, medialogfp);
                                File.Delete(updf);
                                DBControl.AttachPdf(aData.MotconsuId, aData.PatientId, Path.GetFileName(updf), aData.ResDate, folder, Processing.Ini.RubricsId);
                            }

                        }

                    }
                }
                //*/
            }
        }

        class Test
        {
            //Загрузка результатов
            static public void LoadResults()
            {
                XDocument xdoc;

                string[] resfiles = Directory.GetFiles(Ini.ResultFilePath, "*.xml");
                foreach (string filepath in resfiles)
                {
                    xdoc = XDocument.Load(filepath);
                    Console.WriteLine("Загружен файл результатов {0}, обработка файла... ", filepath);

                    string IsDone = xdoc.Element("root").Element("Order").Element("Tests").Element("Test1").Attribute("IsDone") == null ? "" :
                         xdoc.Element("root").Element("Order").Element("Tests").Element("Test1").Attribute("IsDone").Value;
                    if (IsDone == "1")
                    {
                        DataContext db = new DataContext(sqlConnectionStr);
                        Table<Impdata> TImpdata = db.GetTable<Impdata>();

                        XElement PatInfo = xdoc.Element("root").Element("Order").Element("Patient");
                        bool ok = false;
                        string StrPatID;

                        if (PatInfo.Element("MedCardID") != null) StrPatID = PatInfo.Element("MedCardID").Value;
                        else StrPatID = "";

                        string RecordId = xdoc.Element("root").Element("Order").Attribute("OrderID").Value;

                        string[] NameArray = PatInfo.Element("LastName").Value.Split(' ');
                        string[] IOArray;
                        string IO = NameArray.Count() > 1 ? NameArray[1] : "";
                        string FName = "";
                        string MName = "";

                        if (IO.IndexOf('.') > -1 && NameArray.Count() < 3)
                        {
                            IOArray = IO.Split('.');
                            FName = IOArray[0];
                            MName = IOArray.Count() > 1 ? IOArray[1] : "";
                        }

                        PatientName Name = new PatientName
                        {
                            LastName = NameArray[0],
                            FirstName = FName == "" ? (NameArray.Count() > 1 ? NameArray[1] : "") : FName,
                            MiddleName = MName == "" ? (NameArray.Count() > 2 ? NameArray[2] : "") : MName,
                        };

                        if (!DateTime.TryParse(PatInfo.Element("DOB").Value, out DateTime PatBirthDate))
                        {
                            if (PatInfo.Element("DOB").Value.Length == 4 & int.TryParse(PatInfo.Element("DOB").Value, out int BYear))
                                PatBirthDate = new DateTime(BYear, 1, 1);
                            else PatBirthDate = new DateTime(1900, 1, 1);
                        }
                        ok = DBControl.GetPatientID(db, StrPatID, RecordId, out int p_id);
                        if (!ok)
                        {
                            PatientBirthDate BirthDate = new PatientBirthDate
                            {
                                BirthYear = PatBirthDate.Year,
                                BirthMonth = PatBirthDate.Month,
                                BirthDay = PatBirthDate.Day
                            };
                            ok = DBControl.GetPatientID(db, Name, BirthDate, RecordId, out p_id);
                        }

                        if (Ini.TestPatientID != 0) { ok = true; p_id = Ini.TestPatientID; }

                        Console.WriteLine(ok);
                        Console.WriteLine(p_id);

                        string KeyCode = xdoc.Element("root").Element("Order").Attribute("OrderID").Value;

                        bool NewImp = DBControl.GetImpdataID(db, KeyCode, out int ImpdataId);

                        DateTime.TryParseExact(string.Format("{0} {1}", xdoc.Element("root").Element("Header").Element("FileDate").Value,
                                  xdoc.Element("root").Element("Header").Element("FileTime").Value),
                                  new string[] { "yyMMdd HH:mm:ss", "yyMMdd h:mm:ss" }, null, DateTimeStyles.None, out DateTime resdate);

                        if (NewImp)
                        {
                            Console.WriteLine("Создана запиь импорта, id записи {0}", ImpdataId);

                            var insetImpdata = from xe in xdoc.Elements("root")
                                               select new Impdata
                                               {
                                                   ImpdataId = ImpdataId,
                                                   KEYCODE = KeyCode,
                                                   Nom = Name.LastName,
                                                   Prenom = string.Format("{0} {1}", Name.FirstName, Name.MiddleName),
                                                   Date_Naissance = PatBirthDate,
                                                   Date_Consultation = resdate,
                                                   Mesure = Mesure,
                                                   FilialCode = xe.Element("Header").Element("ClinicID").Value
                                               };
                            //Вставка в Impdata
                            TImpdata.InsertAllOnSubmit(insetImpdata);
                            db.SubmitChanges();
                        }

                        //Выборка для вставки в DS_RESTESTS
                        List<RestestData> qRestests = (from xe in xdoc.Elements("root").Elements("Order").Elements("Tests").Descendants().Elements("Result")
                                                       where xe.Element("res") != null
                                                       select new RestestData
                                                       {
                                                           VAL = xe.Element("res").Value,  //GetResultValue(xe.Element("res").Value)
                                                           UNIT = xe.Element("quant") == null ? "" : xe.Element("quant").Value,
                                                           Rescode = xe.Attribute("id").Value,
                                                           ResName = "",//GetResultName(xe.Element("res").Value),
                                                           NORM_TEXT_REC = xe.Element("norm") == null ? "" : xe.Element("norm").Value
                                                       }).ToList();

                        SLS.GetResultValues(qRestests, out List<RestestData> results, out List<RestestData> antibiotics);

                        DBControl.InsertRestestData(db, results, ImpdataId, resdate);

                        //Выборка и вставка антибоиотиков
                        if (antibiotics.Count()>0)
                        {
                            DBControl.InsertRestestData(db, antibiotics, ImpdataId, resdate);

                            //Построение лабораторных справочников: Чувствительность к антибиотикам
                            if (Ini.AntbGroupId != 0)
                            {
                                DBControl.BuildLabDitionary(db, Ini.AntbGroupId, 8);
                            }
                        }

                        if (ok & NewImp) DBControl.UpdateImpdata(db, ImpdataId, p_id);

                        db.Dispose();

                        if (ok)
                        {
                            DBControl.ErrInsertToMotconsu(ImpdataId, p_id, resdate, NewImp);
                            //DBControl.InsetToUserTables(ImpdataId);
                            //Дублирование комментариев в отдельную таблицу
                            DBControl.InsertToEMCTables(ImpdataId);
                            DBControl.UpdateDirAnsw(ImpdataId);
                        }
 
                    }
                }
            }

        }

    }
}
