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

namespace LabComm
{
    public class HelixF
    {
        public static void LoadResults()
        {
            XDocument xdoc = null;
            int p_id = 0;

            //string[] resfiles = Directory.GetFiles(Processing.Ini.ResultFilePath, "*.xml");
            //bool attachPdf = Directory.Exists(Processing.Ini.PdfFilePath) && (Processing.Ini.RubricsId != 0);
            int motconsuId = 0;
            DateTime resdate = DateTime.Now;


            bool attachPdf = Directory.Exists(Processing.Ini.PdfFilePath) && (Processing.Ini.RubricsId != 0);
            string resFolder = @"/%2fOut";
            //string pdfFolder = @"/%2fOutputPDF";
            List<string> filelist;

            if (Processing.Ini.TestMode == 2)
            {
                Console.WriteLine("Test mode is on!");
                filelist = Directory.GetFiles(Processing.Ini.ResultFilePath, "*.xml").ToList();
            }
            else
            {
                try
                {
                    filelist = Files.FTPGetFileList(resFolder, ".xml", false);
                    Console.WriteLine("На серввере {0} файлов для загрузки", filelist.Count);
                }
                catch (Exception ex)
                {
                    filelist = null;
                    Files.WriteLogFile("Ошибка FTP: " + ex.ToString()); Console.WriteLine(ex.Message);
                    Environment.Exit(1);
                }
            }

            foreach (string fileName in filelist)
            {
                //Загрузка файла
                Console.WriteLine("Загрузка результатов из файла {0}", fileName);
                if (Processing.Ini.TestMode == 2)
                {
                    xdoc = XDocument.Load(fileName);
                }
                else
                {
                    try
                    {
                        //fstream = Files.FTPDownload(resFolder + @"/" + fileName);
                        //xdoc = XDocument.Load(fstream);
                        xdoc = Files.FTPDownloadXML(resFolder + @"/" + fileName, false);
                    }
                    catch (Exception ex)
                    {
                        Files.WriteLogFile("Ошибка FTP: " + ex.ToString()); Console.WriteLine(ex.Message);
                        Environment.Exit(1);
                    }
                }

                string[] nameParts = fileName.Split(new char[] { '_', '.' });
                string filialCode = nameParts[nameParts.Length - 2];
                //Опция для нескольких филиалов

                DBControl.GetDepId(filialCode, Processing.Ini.MedecinId, out int depId, out int medDepId);
                Processing.MedDepId = medDepId; Processing.DepId = depId;
                Processing.MedicinsId = Processing.Ini.MedecinId;
                Console.WriteLine("FilialCode {0}\t Meddep {1}\t Dep {2}", filialCode, Processing.MedDepId, Processing.DepId);

                DataContext db = new DataContext(Processing.sqlConnectionStr);

                Table<Impdata> TImpdata = db.GetTable<Impdata>();

                Console.WriteLine("Загрузка результатов из файла {0}", fileName);

                XElement orderX = xdoc.Element("response").Element("Order");

                bool ok = true;

                Processing.PatientName Name = new Processing.PatientName
                {
                    LastName = orderX.Attribute("last_name").Value,
                    FirstName = orderX.Attribute("first_name").Value,
                    MiddleName = orderX.Attribute("middle_name").Value
                };

                DateTime PatBirthDate = DateTime.Parse(orderX.Attribute("birthdate").Value);

                Processing.PatientBirthDate BirthDate = new Processing.PatientBirthDate
                {
                    BirthYear = PatBirthDate.Year,
                    BirthMonth = PatBirthDate.Month,
                    BirthDay = PatBirthDate.Day
                };

                ok = DBControl.GetPatientID(db, Name, BirthDate, fileName, out p_id);


                XElement sampleX = orderX.Element("Sample");

                string KeyCode = orderX.Attribute("lis_order_id").Value;
                string misSampleId = sampleX.Attribute("mis_sample_id").Value;
                string pdffp = string.Format(@"{0}\{1}.pdf", Processing.Ini.ResultFilePath, misSampleId);

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
                    Console.WriteLine("1 - {0}, 2 - {1}", resdate, Processing.Mesure);
                    //Вставка в Impdata
                    if (NewImport)
                    {
                        Processing.ImpData impData = new Processing.ImpData()
                        {
                            ImpdataId = ImpdataId,
                            KEYCODE = KeyCode,
                            Nom = Name.LastName,
                            Prenom = Name.FirstName + " " + Name.MiddleName,
                            Date_Naissance = new DateTime(BirthDate.BirthYear, BirthDate.BirthMonth, BirthDate.BirthDay),
                            Date_Consultation = resdate,
                            Mesure = Processing.Mesure,
                            LabCode = misSampleId
                        };
                        DBControl.InsertImpdata(db, impData);
                    }



                    List<Processing.RestestData> results = new List<Processing.RestestData>();
                    //Комментарии к заявке
                    if (sampleX.Attribute("comment") != null)
                    {
                        int methodId = DBControl.LabMethodCheck(db, "ReqComment");
                        //Создание справочников
                        if (methodId == 0)
                        {
                            Files.WriteLogFile(string.Format("Создание справочников Name {0}, Code {1}, Item_code {2}, Group_ID {3}", "Комментарии к заявке"
                                , "ReqComment", "", Processing.Ini.ResGroupId));
                            DBControl.BuildDitionary(db, "Комментарии к заявке", "ReqComment", "", "", Processing.Ini.ResGroupId, out methodId);
                        }
                        string val = sampleX.Attribute("comment").Value;
                        if (val != string.Empty)
                        {
                            //Компановка результатов
                            results.Add(new Processing.RestestData
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
                        {
                            Files.WriteLogFile(string.Format("Создание справочников Name {0}, Code {1}, Item_code {2}, Group_ID {3}", "Биоматериал", "Biomaterial",
                                    "", Processing.Ini.ResGroupId));
                            DBControl.BuildDitionary(db, "Биоматериал", "Biomaterial", "", "", Processing.Ini.ResGroupId, out methodId);
                        }
                        string val = sampleX.Attribute("Biomaterial").Value;
                        if (val != string.Empty)
                        {
                            //Компановка результатов
                            results.Add(new Processing.RestestData
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
                                {
                                    Files.WriteLogFile(string.Format("Создание справочников Name {0}, Code {1}, Item_code {2}, Group_ID {3}", testX.Attribute("Reported_name").Value, testX.Attribute("Code").Value
                                                     , testX.Attribute("Item_code").Value, Processing.Ini.ResGroupId));
                                    DBControl.BuildDitionary(db, testX.Attribute("Reported_name").Value, testX.Attribute("Code").Value
                                      , testX.Attribute("Item_code").Value, unit, Processing.Ini.ResGroupId, out methodId);
                                }
                                //Компановка результатов
                                string ref_lo = testX.Element("Result").Attribute("ref_lo") != null ? testX.Element("Result").Attribute("ref_lo").Value : string.Empty;
                                string ref_hi = testX.Element("Result").Attribute("ref_hi") != null ? testX.Element("Result").Attribute("ref_hi").Value : string.Empty;
                                results.Add(new Processing.RestestData
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
                                int ResGroupNewId = DBControl.ResGroupCheck(db, testX.Attribute("Item_code").Value, Processing.Ini.ResGroupId);
                                //Создание новой группы исследований
                                if (ResGroupNewId == 0)
                                {
                                    Files.WriteLogFile(string.Format("Создание новой группы Name {0}, Code {1}, Group_ID {2}", testX.Attribute("Reported_name").Value
                                        , testX.Attribute("Item_code").Value, Processing.Ini.ResGroupId));
                                    DBControl.CreateResGroup(db, testX.Attribute("Reported_name").Value, testX.Attribute("Item_code").Value, Processing.Ini.ResGroupId, out ResGroupNewId);
                                }
                                int k = 0;
                                foreach (XElement resultX in testX.Nodes())
                                {
                                    string ncode = testX.Attribute("Item_code").Value + " " + resultX.Attribute("Name").Value;
                                    int methodId = DBControl.LabMethodCheck(db, ncode);
                                    //Создание справочников
                                    string unit = resultX.Attribute("UnitName") != null ? resultX.Attribute("UnitName").Value : "";
                                    k++;
                                    if (methodId == 0)
                                    {
                                        Files.WriteLogFile(string.Format("Создание справочников Name {0}, Code {1}, Item_code {2}, Group_ID {3}", resultX.Attribute("Name").Value,
                                            ncode, testX.Attribute("Item_code").Value, ResGroupNewId));
                                        DBControl.BuildDitionary(db, resultX.Attribute("Name").Value, ncode, string.Format("{0}-{1}", testX.Attribute("Item_code").Value, k)
                                        , unit, ResGroupNewId, out methodId);
                                    }
                                    //Компановка результатов
                                    string ref_lo = resultX.Attribute("ref_lo") != null ? resultX.Attribute("ref_lo").Value : string.Empty;
                                    string ref_hi = resultX.Attribute("ref_hi") != null ? resultX.Attribute("ref_hi").Value : string.Empty;
                                    results.Add(new Processing.RestestData
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
                                {
                                    Files.WriteLogFile(string.Format("Создание справочников Name {0}, Code {1}, Item_code {2}, Group_ID {3}", "Комментарии к исследованию",
                                            "AnaComment", string.Empty, Processing.Ini.CultGroupId));
                                    DBControl.BuildDitionary(db, "Комментарии к исследованию", "AnaComment", "", "", Processing.Ini.CultGroupId, out methodId);
                                }
                                string val = sampleX.Element("TestMicroorganism").Attribute("Comment").Value;
                                if (val != string.Empty)
                                {
                                    //Компановка результатов
                                    results.Add(new Processing.RestestData
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
                                    string v = mic.Attribute("Value").Value == string.Empty ? mic.Attribute("Reported_name").Value : mic.Attribute("Value").Value;
                                    int methodId = DBControl.LabMethodCheck(db, v);
                                    //Создание справочников
                                    if (methodId == 0)
                                    {
                                        Files.WriteLogFile(string.Format("Создание справочников Name {0}, Code {1}, Item_code {2}, Group_ID {3}", v,
                                            v, string.Empty, Processing.Ini.CultGroupId));
                                        DBControl.BuildDitionary(db, v, v
                                          , string.Empty, string.Empty, Processing.Ini.CultGroupId, out methodId);
                                    }
                                    //Компановка результатов
                                    results.Add(new Processing.RestestData
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
                                    {
                                        Files.WriteLogFile(string.Format("Создание справочников Name {0}, Code {1}, Item_code {2}, Group_ID {3}", mic.Attribute("Name").Value, 
                                            ncode, testX.Attribute("Item_code").Value, Processing.Ini.CultGroupId));
                                        DBControl.BuildDitionary(db, mic.Attribute("Name").Value, ncode, testX.Attribute("Item_code").Value
                                        , string.Empty, Processing.Ini.CultGroupId, out methodId);
                                    }
                                    string val = mic.Attribute("Value").Value;
                                    val = val != string.Empty ? val : mic.Attribute("CountValue").Value;
                                    if (val != string.Empty)
                                    {
                                        //Компановка результатов
                                        results.Add(new Processing.RestestData
                                        {
                                            VAL = val,
                                            MethodId = methodId
                                        }
                                            );
                                    }
                                }
                            }
                            Console.WriteLine("!!!!!");
                            //Антибиотики
                            var antibioticRes = from xe in testX.Elements()
                                                where xe.Name == "Antibiotic"
                                                select xe;
                            int item = 0;
                            foreach (var antRes in antibioticRes)
                            {
                                item++;
                                Console.WriteLine(item);
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
                                    {
                                        Files.WriteLogFile(string.Format("Создание справочников Name {0}, Code {1}, Item_code {2}, Group_ID {3}", antName.Attribute("Name").Value, code
                                                     , antRes.Attribute("Item_code").Value, Processing.Ini.AntbGroupId));
                                        DBControl.BuildDitionary(db, antName.Attribute("Name").Value, code
                                              , antRes.Attribute("Item_code").Value, unit, Processing.Ini.AntbGroupId, out methodId);
                                    }
                                    //Компановка результатов
                                    string val = antName.Attribute("Sensitivity") != null
                                                ? string.Format("{0} | {1} {2}", antName.Attribute("Sensitivity").Value, antName.Attribute("Value").Value, unit)
                                                : antName.Attribute("Value").Value;
                                    results.Add(new Processing.RestestData
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
                                    {
                                        Files.WriteLogFile(string.Format("Создание справочников Name {0}, Code {1}, Item_code {2}, Group_ID {3}", bacName.Attribute("Name").Value, code
                                                     , bacRes.Attribute("Item_code").Value, Processing.Ini.BctrGroupId));
                                        DBControl.BuildDitionary(db, bacName.Attribute("Name").Value, code
                                          , bacRes.Attribute("Item_code").Value, unit, Processing.Ini.BctrGroupId, out methodId);
                                    }
                                    //Компановка результатов
                                    string val = bacName.Attribute("Sensitivity") != null
                                                  ? string.Format("{0} | {1} {2}", bacName.Attribute("Sensitivity").Value, bacName.Attribute("Value").Value, unit)
                                                  : bacName.Attribute("Value").Value;
                                    results.Add(new Processing.RestestData
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
                    if (Processing.Ini.TestPatientID != 0) { ok = true; p_id = Processing.Ini.TestPatientID; }

                    //Вставки в DS_RESTESTS
                    DBControl.InsertRestestDataIDM(db, results, ImpdataId, resdate);
                    string pdfFileName = fileName.Replace("xml", "pdf");
                    //string[] pdffilepathes = Directory.GetFiles(pdfFilePath, pdfFileName);

                    if (ok)
                    {
                        DBControl.UpdateImpdata(db, ImpdataId, p_id);
                        //Запись в ЭМК
                        motconsuId = DBControl.InsertToEMCe(ImpdataId, p_id, resdate, NewImport);

                        //Дублирование комментариев в отдельную таблицу
                        DBControl.InsertToEMCTables(ImpdataId);

                        //Дублирование длинных комментариев к параметрам
                        if (Processing.Ini.CommentTable != null && Processing.Ini.CommentField != null)
                            DBControl.ParamsComments(ImpdataId, motconsuId, p_id, Processing.Ini.CommentTable, Processing.Ini.CommentField);

                        //Обработка pdf файла и прикрепление к ЭМК
                        if (Processing.Ini.TestMode == 0)
                        {
                            //Обработка pdf файла и прикрепление к ЭМК
                            try
                            {
                                if (attachPdf && motconsuId != 0)
                                {
                                    Console.WriteLine("Работа с pdf файлом");
                                    string folder = string.Format("{0}-{1}/{2}", resdate.Year, Convert.ToString(resdate.Month).PadLeft(2, '0'), p_id);
                                    string pdfDirectory = string.Format("{0}/{1}", Processing.Ini.PdfFilePath, folder);
                                    string medialogfp = string.Format("{0}/{1}", pdfDirectory, pdfFileName);

                                    if (!Directory.Exists(pdfDirectory))
                                    {
                                        DirectoryInfo dir = Directory.CreateDirectory(pdfDirectory);
                                    }

                                    //Копирование pdf на сервер Медиалога
                                    if (File.Exists(medialogfp))
                                    {
                                        File.Delete(medialogfp);
                                        DBControl.DeletePdfRefrense(p_id, pdfFileName);
                                    }
                                    
                                    if (Files.FTPFolderInfo(resFolder, false).IndexOf(pdfFileName) != -1)
                                    {
                                        Console.WriteLine("pdf файл найден");
                                        //Files.SaveStreamToFile(medialogfp, pdfstream);
                                        Files.FTPDownloadFile(resFolder + @"/" + pdfFileName, medialogfp, false);
                                        DBControl.AttachPdf(motconsuId, p_id, pdfFileName, resdate, folder, Processing.Ini.RubricsId);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Не возможно обработать pdf файл");
                                    if (Directory.Exists(Processing.Ini.ArchiveFilePath))
                                    {
                                        string pdfarchivefp = string.Format("{0}/{1}", Processing.Ini.ArchiveFilePath, pdfFileName);
                                        if (File.Exists(pdfarchivefp)) File.Delete(pdfarchivefp);
                                        //Files.SaveStreamToFile(pdfarchivefp, pdfstream);
                                        Files.FTPDownloadFile(resFolder + @"/" + pdfFileName, pdfarchivefp, false);
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
                    else
                    {
                        if (Directory.Exists(Processing.Ini.UnboundPdfFilePath))
                        {
                            string unbound = string.Format("{0}/{1}", Processing.Ini.UnboundPdfFilePath, Path.GetFileName(fileName));
                            if (Files.FTPFolderInfo(resFolder, false).IndexOf(pdfFileName) != -1)
                            {
                                unbound = string.Format("{0}/{1}", Processing.Ini.UnboundPdfFilePath, Path.GetFileName(pdffp));
                                if (File.Exists(unbound)) File.Delete(unbound);
                                Files.FTPDownloadFile(resFolder + @"/" + pdfFileName, unbound, false);
                                Files.FTPDelete(resFolder + @"/" + pdfFileName, false);
                            }
                        }

                    }
                    db.Dispose();

                    if (Directory.Exists(Processing.Ini.ArchiveFilePath))
                    {
                        string archivefp = string.Format("{0}/{1}", Processing.Ini.ArchiveFilePath, fileName);
                        if (File.Exists(archivefp)) File.Delete(archivefp);
                        if (Processing.Ini.TestMode == 0)
                        {
                            Files.FTPDownloadFile(resFolder + @"/" + fileName, archivefp, false);
                            //Files.SaveStreamToFile(archivefp, fstream);
                        }
                    }

                    if (Processing.Ini.TestMode == 0)
                    {
                        //Удаление файлов с FTP-сервера
                        if (Processing.Ini.DeleteDownloaded)
                        {
                            Files.FTPDelete(resFolder + @"/" + fileName, false);
                            if (Files.FTPFolderInfo(resFolder, false).IndexOf(pdfFileName) != -1)  Files.FTPDelete(resFolder + @"/" + pdfFileName, false);
                        }


                    }

                    /*
                    //Отправка файла во временную директорию или удаление
                    if (Directory.Exists(Processing.Ini.ArchiveFilePath))
                    {
                        string archivefp = string.Format("{0}/{1}", Processing.Ini.ArchiveFilePath, Path.GetFileName(filepath));
                        if (File.Exists(archivefp)) File.Delete(archivefp);
                        File.Copy(filepath, archivefp);
                        File.Delete(filepath);
                    }
                    else
                    {
                        if (Processing.Ini.DeleteDownloaded)
                        {
                            File.Delete(filepath);
                        }
                    }

                    if (Directory.Exists(Processing.Ini.UnboundPdfFilePath))
                    {
                        string unbound = string.Format("{0}/{1}", Processing.Ini.UnboundPdfFilePath, Path.GetFileName(filepath));
                        if (File.Exists(pdffp))
                        {
                            unbound = string.Format("{0}/{1}", Processing.Ini.UnboundPdfFilePath, Path.GetFileName(pdffp));
                            if (File.Exists(unbound)) File.Delete(unbound);
                            File.Copy(pdffp, unbound);
                            File.Delete(pdffp);
                        }
                    }
                    else
                    {
                        if (File.Exists(pdffp)) File.Delete(pdffp);
                    }
                    */
                }
                else
                {
                    Console.WriteLine("Статус образца: Отменен");
                }
            }

            //Повторная обработка неприкрепленных pdf файлов
            if (Directory.Exists(Processing.Ini.UnboundPdfFilePath))
            {
                string[] unPdffiles = Directory.GetFiles(Processing.Ini.UnboundPdfFilePath, "*.pdf");
                foreach (string updf in unPdffiles)
                {
                    string fNamePart = Path.GetFileName(updf).Split(new char[] { '.' })[0];
                    Console.WriteLine(updf);
                    Processing.AttachData aData = DBControl.GetAttachDataS(fNamePart);

                    if (attachPdf && aData != null)
                    {
                        string folder = string.Format("{0}-{1}/{2}", aData.ResDate.Year, Convert.ToString(aData.ResDate.Month).PadLeft(2, '0'), aData.PatientId);
                        string pdfDirectory = string.Format("{0}/{1}", Processing.Ini.PdfFilePath, folder);

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
}
