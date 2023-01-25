using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Timers;
using System.Threading;
using System.Text.RegularExpressions;

namespace LabComm
{
    public class Invitro 
    {
        public class Products
        {
            public string Code;
            public string Label;
            public string ShortLabel;
            public string ExternalId;
            public string OptionSet;
            public string BiomId;
            public XElement AuxiliaryInfoIds;
            public string GroupName;
            public string SubgroupName;
        }

        public class Auxilary
        {
            public string ExternalId;
            public string Name;
            public bool IsRequired;
            public string Min;
            public string Max;
            public string Unit;
            public string ValType;
            public string Variants;
        }

        public class AuxProfile
        {
            public string AuxExternalId;
            public string ProfileCode;
            public string BiomatCode;
        }

        public class StickerInfo
        {
            public string BarCodeNumber;
            public string BiomaterialId;
            public string ContainerId;
            public string ZPLText;
        }

        public class StickerProfile
        {
            public string ProductId;
            public string INZ;
        }

        public class OrderCoverLetter
        {
            public string OrderId;
            public string PdfContent;
        }

        public class TestTube
        {
            public string Code;
            public string Name;
            public string Description;
            public string ExternalId;
            public string Type;
            public string TypeCode;
        }

        public class AuxilaryInfo
        {
            public string AuxExtCode;
            public string AuxTableName;
            public string AuxFieldName;
            public int? MotconsuId;
        }

        public class AuilaryData
        {
            public string AuxExtCode;
            public string AuxValue;
        }

        public class ExtendedDict
        {
            public int Id;
            public string Name;
            public string ExtendedId;
            public int RefId;
        }

        public class ExtendedProduct
        {
            public string ExternalId;
            public string OptionSet;
            public string BiomaterialId;
            public string ContainerId;
            public int FormType;
        }

        public static void LoadDictionaries()
        {
            XNamespace xsd = "http://www.w3.org/2001/XMLSchema";
            XNamespace sxsi = "http://www.w3.org/2001/XMLSchema-instance";
            Console.WriteLine("point1.1");
            List<ProfileBiomat> profileBiomat = new List<ProfileBiomat>();
            List<AuxProfile> auxProfiles = new List<AuxProfile>();
            try
            {
                XDocument Response;
                string errMessage;
                string getReq="";
                //getReq = string.Format("{0}/{1}/{2}", Processing.Ini.url, "GetExtendedProducts", Processing.Ini.token);
                // получение xml по http
                //Message.Invitro getExtProduct = new Message.Invitro(getReq); //GetInfo GetExtendedInfo GetProducts
                Response = XDocument.Load(@"C:\Users\79529\Documents\DATA\Invitro\ExtendedProducts.xml");

                if (1==1) // выполнение метода getExtProduct.GetRequest(out Response, out errMessage)
                {
                    if (Processing.Ini.WriteInput) // запись в xml файле
                    {
                        Response.Save(string.Format("{0}/{1}.xml", Files.InputDir, "ExtendedProducts"));
                        Console.WriteLine("Ответ {0}.xml помещен в папку {1}", "ExtendedProducts", Files.InputDir);
                    }
                    // считывание xml в список
                    var extendedProducts = (from a in Response.Element("ArrayOfExtendedProduct").Elements("ExtendedProduct").Elements("ExtendedBiomaterialOptionSets")
                                        .Elements("ExtendedBiomaterialOptionSet").Elements("BiomaterialOptions").Elements("ExtendedBiomaterialOption")
                                        .Elements("ContainerInfos").Elements("ExtendedContainerInfo")
                                    select new ExtendedProduct
                                    {
                                        ExternalId = a.Parent.Parent.Parent.Parent.Parent.Parent.Element("Id").Value,
                                        OptionSet = a.Parent.Parent.Parent.Parent.Element("Id").Value,
                                        BiomaterialId = a.Parent.Parent.Element("BiomaterialId").Value,
                                        ContainerId = a.Element("ContainerId").Value,
                                        FormType = Convert.ToInt32(a.Element("FormType").Value)
                                    }).ToList();
                    DBControl.InsertExtendedProfiles(extendedProducts); // вставка в БД
                }

                getReq = string.Format("{0}/{1}/{2}", Processing.Ini.url, "GetProducts", Processing.Ini.token);
                Response = XDocument.Load(@"C:\Users\79529\Documents\DATA\Invitro\Products.xml");
                //Message.Invitro getProduct = new Message.Invitro(getReq);//GetInfo GetExtendedInfo GetProducts
                if (1==1) //getProduct.GetRequest(out Response, out errMessage)
                {
                    if (Processing.Ini.WriteInput)
                    {
                        Response.Save(string.Format("{0}/{1}.xml", Files.InputDir, "Products"));
                        Console.WriteLine("Ответ {0}.xml помещен в папку {1}", "Products", Files.InputDir);
                    }
                    var products = (from p in Response.Element("ArrayOfInternalProduct").Elements("InternalProduct").Elements("BiomaterialOptionSets")
                   .Elements("BiomaterialOptionSet").Elements("BiomaterialOptions").Elements("BiomaterialOption")
                                    select new Products
                                    {
                                        ExternalId = p.Parent.Parent.Parent.Parent.Element("Id").Value,
                                        Code = p.Parent.Parent.Parent.Parent.Element("Code").Value,
                                        Label = p.Parent.Parent.Parent.Parent.Element("Name").Value,
                                        ShortLabel = p.Parent.Parent.Parent.Parent.Element("ShortName").Value,
                                        OptionSet = p.Parent.Parent.Element("Id").Value,
                                        BiomId = p.Element("BiomaterialId").Value,
                                        AuxiliaryInfoIds = p.Element("AuxiliaryInfoIds"),
                                        GroupName = p.Parent.Parent.Parent.Parent.Element("GroupName") != null ? p.Parent.Parent.Parent.Parent.Element("GroupName").Value : null,
                                        SubgroupName = p.Parent.Parent.Parent.Parent.Element("SubgroupName") != null ? p.Parent.Parent.Parent.Parent.Element("SubgroupName").Value : null
                                    });
                    var profiles = products.GroupBy(g => new { g.Code, g.Label, g.ShortLabel, g.ExternalId, g.GroupName, g.SubgroupName })
                        .Select(p => new Profile { Code = p.Key.Code, Label = p.Key.Label, ExternalId = p.Key.ExternalId, ShortLabel = p.Key.ShortLabel
                                , GroupName = p.Key.GroupName, SubgroupName = p.Key.SubgroupName })
                        .ToList();
                    DBControl.InsertProfiles(profiles, Processing.Ini.OrgId);
                    profileBiomat = products.Select(p => new ProfileBiomat { ProfileCode = p.Code, BiomatCode = p.BiomId, OptionSet = p.OptionSet, ProfileExtId = p.ExternalId }).ToList();

                    foreach (var p in products)
                    {
                        if (p.AuxiliaryInfoIds.Descendants("AuxiliaryInfoId").Any())
                        {
                            foreach (var a in p.AuxiliaryInfoIds.Elements())
                            {
                                auxProfiles.Add(new AuxProfile
                                {
                                    ProfileCode = p.Code,
                                    BiomatCode = p.BiomId,
                                    AuxExternalId = a.Value
                                });
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine(errMessage);
                    Files.WriteLogFile(errMessage);
                }
                Console.WriteLine("Token: {0}", Processing.Ini.token);
                getReq = string.Format("{0}/{1}/{2}", Processing.Ini.url, "GetInfo", Processing.Ini.token);
                Response = XDocument.Load(@"C:\Users\79529\Documents\DATA\Invitro\InfoLab.xml");
                //Message.Invitro getInfo = new Message.Invitro(getReq); //GetInfo GetExtendedInfo GetProducts
                if (1==1)//getInfo.GetRequest(out Response, out errMessage)
                {
                    if (Processing.Ini.WriteInput)
                    {
                        Response.Save(string.Format("{0}/{1}.xml", Files.InputDir, "InfoLab"));
                        Console.WriteLine("Ответ {0}.xml помещен в папку {1}", "InfoLab", Files.InputDir);
                    }
                    var auxilary = (from a in Response.Element("Data").Element("AuxiliaryInfos").Elements("AuxiliaryInfo")
                                    select new Auxilary
                                    {
                                        ExternalId = a.Element("Id").Value,
                                        Name = a.Element("Name").Value,
                                        IsRequired = a.Element("IsRequired").Value == "true" ? true : false,
                                        Min = GetValue(a.Element("Min")),
                                        Max = GetValue(a.Element("Max")),
                                        Unit = GetValue(a.Element("Unit")),
                                        ValType = GetValue(a.Element("ValueType")),
                                        Variants = a.Descendants("string").Any() ? GetVariants(a.Element("Variants")) : string.Empty   //string.Join(Environment.NewLine, a.Element("Variants").Elements(""))
                                    }).ToList();
                    DBControl.InsertAuxilary(auxilary);

                    var biomat = (from b in Response.Element("Data").Element("Biomaterials").Elements("Biomaterial")
                                  select new Biomat
                                  {
                                      Code = b.Element("Id").Value,
                                      Label = b.Element("Name").Value
                                  }).ToList();
                    DBControl.InsertBiomaterial(biomat, Processing.Ini.OrgId);

                    var testTubes = (from c in Response.Element("Data").Element("TestTubes").Elements("TestTube")
                                     select new TestTube
                                     {
                                         Code = c.Element("Code").Value,
                                         Name = c.Element("Name").Value,
                                         Description = c.Element("ContainerDescription") != null ? c.Element("ContainerDescription").Value : string.Empty,
                                         ExternalId = c.Element("Id").Value,
                                         Type = c.Element("Type").Value,
                                         TypeCode = c.Element("TextId").Value
                                     }).ToList();
                    DBControl.InsertContainer(testTubes, Processing.Ini.OrgId);
                }
                else
                {
                    Console.WriteLine(errMessage);
                    Files.WriteLogFile(errMessage);
                }



                /*
                getReq = string.Format("{0}/{1}/{2}", Processing.Ini.url, "GetExtendedInfo", Processing.Ini.token);
                Message.Invitro getExtInfo = new Message.Invitro(getReq); //GetInfo GetExtendedInfo GetProducts
                if (getExtInfo.GetRequest(out Response, out errMessage))
                {
                    List<ExtendedDict> dictData;

                    dictData = (from b in Response.Element("Data").Element("DocumentTypes").Elements("DocumentType")
                                  select new ExtendedDict
                                  {
                                      Id = Convert.ToInt32(b.Element("Id").Value),
                                      Name = b.Element("Name").Value
                                  }).ToList();
                    DBControl.InsertDocType(dictData);

                    dictData = (from b in Response.Element("Data").Element("AddressTypes").Elements("AddressType")
                                select new ExtendedDict
                                {
                                    Id = Convert.ToInt32(b.Element("Id").Value),
                                    Name = b.Element("Name").Value
                                }).ToList();
                    DBControl.InsertAdresType(dictData);

                    dictData = (from b in Response.Element("Data").Element("Countries").Elements("Country")
                                select new ExtendedDict
                                {
                                    Id = Convert.ToInt32(b.Element("Id").Value),
                                    Name = b.Element("Name").Value
                                }).ToList();
                    DBControl.InsertCountry(dictData);
                    //foreach(var b in Response.Element("Data").Element("Regions").Elements("Region")) 
                        //Console.WriteLine("0 = {0}\t1 = {1}\t2 = {2}\t3 = {3} ", b.Element("Id").Value, b.Element("Name").Value, b.Element("ExtendedId").Value, b.Element("Country").Value);

                    dictData = (from b in Response.Element("Data").Element("Regions").Elements("Region")
                                where b.Element("Name") != null
                                select new ExtendedDict
                                {
                                    Id = Convert.ToInt32(b.Element("Id").Value),
                                    Name = b.Element("Name").Value,
                                    ExtendedId = b.Element("ExtendedId").Value,
                                    RefId = Convert.ToInt32(b.Element("Country").Value)
                                }).ToList();
                    DBControl.InsertRegion(dictData);

                    dictData = (from b in Response.Element("Data").Element("Cities").Elements("City")
                                select new ExtendedDict
                                {
                                    Id = Convert.ToInt32(b.Element("Id").Value),
                                    Name = b.Element("Name").Value,
                                    ExtendedId = b.Element("ExtendedId").Value,
                                    RefId = Convert.ToInt32(b.Element("RegionId").Value)
                                }).ToList();
                    DBControl.InsertCity(dictData);
                }
                else
                {
                    Console.WriteLine(errMessage);
                    Files.WriteLogFile(errMessage);
                }
                //*/
                DBControl.InsertBiomatProfile(profileBiomat, Processing.Ini.OrgId);
                /*DBControl.DeleteBiomatProfile(profileBiomat, Processing.Ini.OrgId);*/
                DBControl.InsertAuxilaryProfile(auxProfiles, Processing.Ini.OrgId);
                DBControl.UpdateFormType();
                DBControl.UpdateFormTypelog();

                string GetValue(XElement xElement)
                {
                    if (xElement == null)
                        return string.Empty;
                    else return xElement.Value;
                }

                string GetVariants(XElement xElement)
                {
                    string res = "";
                    int i = 0;
                    foreach (var e in xElement.Elements())
                    {
                        res += (i == 0 ? e.Value : Environment.NewLine + e.Value);
                        i++;
                    }
                    return res;
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Files.WriteLogFile(ex.ToString());
                Thread.Sleep(5000);
                Environment.Exit(1);
            }
            Console.WriteLine("point1.2");
        }

        //Запуск создания заказов
        static public void OrderCreate(object obj, ElapsedEventArgs e)
        {
            try
            {
                int? state = DBControl.bcChanged();
                if (state != null)
                {
                    Processing.bcChangeTimer.Enabled = false;
                    Console.WriteLine("Изменения в таблице заказов... {0}", state);
                    if (state == 1)
                    {
                        Console.WriteLine("Отправка заказов");
                        Thread.Sleep(500);
                        DBControl.UpdateChangedOrder(Convert.ToInt32(state));
                        UplodOrders();
                    }
                    if (state == -1)
                    {
                        Console.WriteLine("Удаление заказов");
                        Thread.Sleep(100);
                        DBControl.UpdateChangedOrder(Convert.ToInt32(state));
                        DeleteOrders();
                    }
                    Processing.bcChangeTimer.Enabled = true;
                }
            }
            catch(Exception ex)
            {
                Files.WriteLogFile(ex.ToString());
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Processing.bcChangeTimer.Enabled = true;
            }
        }

        //Удаление заказов
        static public void OrderDelete(object obj, ElapsedEventArgs e)
        {
            if (DBControl.delChanged(out int checksum))
            {
                 Console.WriteLine("Обнаружен признак удаления заказа!");
                DeleteOrders();
            }
        }

        static public void DeleteOrders()
        {
            List<string> ordersToDelete = DBControl.GetOrdersToDelete();
            foreach (string extId in ordersToDelete)
            {
                string getReq = string.Format("{0}/{1}/{2}", Processing.Ini.url, "RemoveOrder", extId);
                XDocument Response;
                string errMessage;
                //DBControl.DeleteOrder(extId);
                //Files.WriteLogFile(getReq);
                Message.Invitro deleteOrder = new Message.Invitro(getReq, "DELETE");//GetInfo GetExtendedInfo GetProducts
                ///*
                if (deleteOrder.Request(out Response, out errMessage))
                {
                    Files.WriteLogFile(Response.ToString());
                    DBControl.DeleteOrder(extId);
                    Console.WriteLine("Заказ {0} удален", extId);
                }
                else
                {
                    Files.WriteLogFile(errMessage);
                    try
                    {
                        XDocument emsg = XDocument.Parse(errMessage);
                        XNamespace xns = "http://schemas.datacontract.org/2004/07/Invitro.Integration.Scape.Contracts";
                        string eOrder = emsg.Element(xns + "ErrorMessage").Element(xns + "UserFriendlyMessage").Value;
                        Console.WriteLine(eOrder);
                    }
                    catch
                    {
                        Console.WriteLine(errMessage);
                    }
                }
                //*/
            }
        }

        //
        static public void UplodOrders()
        {
            CultureInfo culture = new CultureInfo("ru-RU"); //формат преобразования даты
            List<Request> RequestInfoList = new List<Request>();
            //Files.GoFurther();
            try
            {
                DBControl.GetOrdersListV2(2, 0, out RequestInfoList);
                //Формирование заказов
                if (RequestInfoList.Count != 0)
                {
                    Console.WriteLine("Готово к отправке {0} заказов", RequestInfoList.Count);
                    Console.WriteLine("Создание и выгрузка готовых заказов..."); 
                    ///*
                    foreach (Request qRequest in RequestInfoList)
                    {
                        //Получение данных заказа из базы данных
                        DBControl.GetOrderInfoInvitro(qRequest.ID, out PatientInfo qPatient, out List<OrderInvitro> qOrder);

                        var requests = qOrder.GroupBy(g => new { g.FilialId, g.RequestId, g.RequestDate, g.RequestCode });//  .Select(x => x.Key.FilialId);

                        foreach (var request in requests)
                        {
                            //DateTime ReqDateTime = DateTime.SpecifyKind(request.Key.RequestDate, DateTimeKind.Local);

                            //string Token = DBControl.GetToken(request.Key.FilialId);
                            FilialInfo filialInfo = DBControl.GetFilialInfo(request.Key.FilialId);
                            string Token = filialInfo.Token;
                            Console.WriteLine("Token: {0}", Token);
                            var products = request.GroupBy(g => new { g.ProductExternalId }).Select(x => x.Key);
                            var bmoptions = request.GroupBy(g => new { g.ProductExternalId, g.BiomaterialOption }).Select(x => x.Key);
                            var biomaterials = request.GroupBy(g => new { g.BiomaterialOption, g.BiomaterialCode }).Select(x => x.Key);
                            var auxinfo = request.Where(x => x.MotconsuId != null && x.AuxFieldName != null)
                                       //.GroupBy(g => new AuxilaryInfo { AuxExtCode = g.AuxExtCode, AuxTableName = g.AuxTableName, AuxFieldName = g.AuxFieldName, MotconsuId = g.MotconsuId })
                                       .GroupBy(g => new { g.AuxExtCode,  g.AuxTableName,  g.AuxFieldName,  g.MotconsuId })
                                        //.Select(x => x.Key).ToList();
                                       .Select(x => new AuxilaryInfo { AuxExtCode = x.Key.AuxExtCode, AuxTableName = x.Key.AuxTableName, AuxFieldName = x.Key.AuxFieldName, MotconsuId = x.Key.MotconsuId }).ToList();

                            var auxdata = DBControl.GetAuxData(auxinfo);
                            int? motconsuId = request.Select(x => x.MotconsuId).FirstOrDefault();
                            string reqCode = request.Select(x => x.RequestCode).FirstOrDefault();
                            foreach (var inf in auxinfo) Console.WriteLine(string.Format("{0} {1} {2}",inf.AuxExtCode, inf.MotconsuId, inf.AuxFieldName));
                            foreach (var aux in auxdata) Console.WriteLine(aux.AuxExtCode + ' ' + (DateTime.TryParse(aux.AuxValue, out DateTime dat) ? dat.ToString("dd.MM.yy"): aux.AuxValue));
                            //Создание тела XML сообщения для создания заказа в ЛИС 

                            XElement xpatient = new XElement("Patient",
                                                   new XElement("ExternalId", qPatient.ID),
                                                   new XElement("LastName", qPatient.LastName),
                                                   new XElement("FirstName", qPatient.FirstName),
                                                   new XElement("MiddleName", qPatient.MiddleName),
                                                   new XElement("BirthDate", qPatient.BirthDate),
                                                   new XElement("Sex", qPatient.Gender == 0 ? "M" : qPatient.Gender == 1 ? "F" : "-")
                                                  );
                            if ((qPatient.addrReg ?? string.Empty) != string.Empty) xpatient.Add(new XElement("Adress", qPatient.addrReg));
                            if ((qPatient.phone ?? string.Empty) != string.Empty) xpatient.Add(new XElement("PhoneNumber", qPatient.phone));
                            if ((qPatient.workPlace ?? string.Empty) != string.Empty) xpatient.Add(new XElement("WorkPlace", qPatient.workPlace));
                            if ((qPatient.post ?? string.Empty) != string.Empty) xpatient.Add(new XElement("Post", qPatient.post));
                            if ((qPatient.snils ?? string.Empty) != string.Empty) xpatient.Add(new XElement("SNILS", qPatient.snils));

                            XElement xadres = new XElement("Address",
                                                new XElement("Type", qPatient.adresType == 0 ? 4 : qPatient.adresType));
                            if ((qPatient.regionName ?? string.Empty) != string.Empty) xadres.Add(new XElement("Subject", qPatient.regionName));
                            if ((qPatient.postalCode ?? string.Empty) != string.Empty) xadres.Add(new XElement("PostalCode", qPatient.postalCode));
                            if ((qPatient.countryName ?? string.Empty) != string.Empty) xadres.Add(new XElement("Country", qPatient.countryName));
                            if ((qPatient.district ?? string.Empty) != string.Empty) xadres.Add(new XElement("District", qPatient.district));
                            if ((qPatient.town ?? string.Empty) != string.Empty) xadres.Add(new XElement("Locality", qPatient.town));
                            if ((qPatient.streetName ?? string.Empty) != string.Empty) xadres.Add(new XElement("Street", qPatient.streetName));
                            if ((qPatient.house ?? string.Empty) != string.Empty) xadres.Add(new XElement("House", qPatient.house));
                            if ((qPatient.building ?? string.Empty) != string.Empty) xadres.Add(new XElement("Body", qPatient.building));
                            if ((qPatient.appartament ?? string.Empty) != string.Empty) xadres.Add(new XElement("Flat", qPatient.appartament));

                            XElement xoffice = new XElement("Office",
                                                   new XElement("ClientId", filialInfo.ClientId));

                            XElement xdocument = new XElement("Document",
                                                   new XElement("Type", qPatient.docTypeId));
                            if ((qPatient.docSerNumber ?? string.Empty) != string.Empty) xdocument.Add(new XElement("Series", qPatient.docSerNumber));
                            if ((qPatient.docNumber ?? string.Empty) != string.Empty) xdocument.Add(new XElement("Number", qPatient.docNumber));
                            if ((qPatient.issueOrgName ?? string.Empty) != string.Empty) xdocument.Add(new XElement("Issued", qPatient.issueOrgName));
                            if ((qPatient.issueDate) != new DateTime()) xdocument.Add(new XElement("DateOfIssue", qPatient.issueDate));

                            XElement xproducts = new XElement("Products",
                                                  from product in products
                                                  select new XElement("Product",
                                                                new XElement("ProductId", product.ProductExternalId),
                                                                from bmoption in bmoptions
                                                                where product.ProductExternalId == bmoption.ProductExternalId
                                                                select new XElement("BiomaterialOptions",
                                                                            from biomaterial in biomaterials
                                                                            where bmoption.BiomaterialOption == biomaterial.BiomaterialOption
                                                                            select new XElement("Option",
                                                                                   new XElement("Id", biomaterial.BiomaterialOption),
                                                                                   new XElement("BiomaterialId", biomaterial.BiomaterialCode)
                                                                                                )
                                                                                    )     //BiomaterialOptions
                                                                     )   //Product
                                                    );   //Products

                            XDocument messagebody = new XDocument(
                              new XElement("Order", new XAttribute(XNamespace.Xmlns + "xsi", "http://www.w3.org/2001/XMLSchema-instance")
                                        , new XAttribute(XNamespace.Xmlns + "xsd", "http://www.w3.org/2001/XMLSchema")
                             , new XElement("Token", Token)
                             , new XElement("ExternalId", request.Key.RequestId)
                             , new XElement("BiomaterialDate", DateTime.SpecifyKind(request.Key.RequestDate, DateTimeKind.Local))
                             , new XElement("SendToEpgu", "false")
                             , new XElement(xpatient)
                             , new XElement(xproducts)
                             )   //Order
                     );  //XDocument 
                            XElement refElement;
                            refElement = messagebody.Element("Order").Element("Patient");
                            refElement.AddAfterSelf(xadres);

                            refElement = messagebody.Element("Order");
                            if (qPatient.Email != null || qPatient.phone != null)
                            {
                                XElement xDeliveries = new XElement("Deliveries");
                                if (qPatient.Email != null)
                                    xDeliveries.Add(new XElement("Delivery",
                                                            new XElement("Type", "Email"),
                                                            new XElement("Value", qPatient.Email)
                                                         ));
                                if (qPatient.phone != null)
                                    xDeliveries.Add(new XElement("Delivery",
                                                            new XElement("Type", "SmsNotification"),
                                                            new XElement("Value", new Regex("[^0-9 ]").Replace(qPatient.phone, string.Empty))
                                                         ));
                                refElement.Add(xDeliveries);
                            }

                            if (auxdata != null)
                            {
                                XElement xAuxdata = new XElement("AuxiliaryInfoValues",
                                                        from aux in auxdata
                                                        select new XElement("AuxiliaryInfoValue",
                                                                    new XElement("AuxiliaryInfoId", aux.AuxExtCode),
                                                                    new XElement("Value", DateTime.TryParse(aux.AuxValue, out DateTime dat) ? dat.ToString("dd.MM.yy") : aux.AuxValue)
                                                               ) //AuxiliaryInfoValue
                                                        ); //AuxiliaryInfoValues

                                refElement.Add(xAuxdata);
                            }

                            if (Processing.Ini.WriteOutput)
                            {
                                messagebody.Save(string.Format("{0}/{1}.xml", Files.OutputDir, request.Key.RequestCode));
                                Console.WriteLine("Создан фaйл заказа {0}.xml и помещен в папку {1}", request.Key.RequestCode, Files.OutputDir);
                            }

                            if (Processing.Ini.TestMode == 0)
                            {
                                ///*
                                Message.Invitro req = new Message.Invitro();
                                if (req.RegisterOrder(messagebody, out XDocument Response, out string errMessage))
                                {
                                    if (Processing.Ini.WriteInput)
                                    {
                                        Response.Save(string.Format("{0}/{1}{2}.xml", Files.InputDir, "Response", request.Key.RequestCode));
                                        Console.WriteLine("Ответ {0}.xml помещен в папку {1}", "Response", Files.InputDir);
                                    }
                                    Encoding enc = Encoding.GetEncoding(866);
                                    List<StickerInfo> stickerList = (from xe in Response.Elements("OrderResponse").Elements("OrderTubes").Elements("OrderTube")
                                                                     select new StickerInfo
                                                                     {
                                                                         BarCodeNumber = xe.Element("LaboratoryNumber").Value,
                                                                         BiomaterialId = xe.Element("BiomaterialId").Value,
                                                                         ContainerId = xe.Element("ContainerId").Value,
                                                                         ZPLText = enc.GetString(Convert.FromBase64String(xe.Element("StickerCodeBase64").Value))
                                                                     }).ToList();
                                    /*
                                    List<StickerProfile> stickerProfiles = (from xe in Response.Elements("OrderResponse").Elements("ProductsInfos").Elements("ProductsInfo").Elements("Inzs").Elements("int")
                                                                            select new StickerProfile
                                                            {
                                                                ProductId = xe.Parent.Parent.Element("ProductId").Value,
                                                                INZ = xe.Value,
                                                            }).ToList();
                                    
                                    foreach (var stp in stickerProfiles){ Console.WriteLine("{0}\t{1}", stp.ProductId, stp.INZ); }
                                    */
                                    string OrderINZ = Response.Elements("OrderResponse").Elements("ProductsInfos").Elements("ProductsInfo").Elements("Inzs").Elements("int").FirstOrDefault().Value;
                                    string OrderExtId = Response.Elements("OrderResponse").Elements("OrderId").FirstOrDefault().Value;
                                    DBControl.UpdateOrderState(request.Key.RequestId, 3, OrderINZ, OrderExtId);
                                    DBControl.InsertBarcodes(request.Key.RequestId, stickerList);
                                    DBControl.UpdatePatdirecContInfo(request.Key.RequestId);
                                    
                                    if (motconsuId != null && Directory.Exists(Processing.Ini.PdfFilePath) && (Processing.Ini.PatdirRubricsId != 0))
                                    {
                                        Console.WriteLine("Запись pdf направительного бланка ");
                                        DateTime now = DateTime.Now;
                                        string pdfFileName = string.Format("{0}.{1}", reqCode, "pdf");
                                        string folder = string.Format("{0}-{1}/{2}", now.Year, Convert.ToString(now.Month).PadLeft(2, '0'), qPatient.ID);
                                        string pdfDirectory = string.Format("{0}/{1}", Processing.Ini.PdfFilePath, folder);
                                        //string medialogfp = string.Format("{0}/{1}", pdfDirectory, pdfFileName);
                                        string attachFilePath = string.Format("{0}/{1}", pdfDirectory, pdfFileName);
                                        if (!Directory.Exists(pdfDirectory)) 
                                            { DirectoryInfo dir = Directory.CreateDirectory(pdfDirectory); }
                                        string CoverLetterContent = Response.Element("OrderResponse").Element("CoverLetters").Element("CoverLetter").Element("ContentBase64").Value;
                                        File.WriteAllBytes(attachFilePath, Convert.FromBase64String(CoverLetterContent));
                                        DBControl.AttachPdfV2(Convert.ToInt32(motconsuId), qPatient.ID, pdfFileName, now, folder, Processing.Ini.PatdirRubricsId);
                                    }
                                }
                                else
                                {
                                    Files.WriteLogFile(errMessage);
                                    try
                                    {
                                        XDocument emsg = XDocument.Parse(errMessage);
                                        XNamespace xns = "http://schemas.datacontract.org/2004/07/Invitro.Integration.Scape.Contracts";
                                        string eOrder = emsg.Element(xns + "ErrorMessage").Element(xns + "UserFriendlyMessage").Value;
                                        DBControl.UpdateOrderStateError(request.Key.RequestId, 4, "error", eOrder);
                                        Console.WriteLine(eOrder);
                                    }
                                    catch
                                    {
                                        DBControl.UpdateOrderStateError(request.Key.RequestId, 4, "error", errMessage);
                                        Console.WriteLine(errMessage);
                                    }
                                }
                                //*/
                            }
                            //DBControl.UpdateOrderState(request.Key.RequestId, 1, "test01", "test01");
                            if (Processing.Ini.TestMode == 2)
                            {
                                XDocument xdoc;
                                xdoc = XDocument.Load(@"D:\data\TestOrders\Response1905070014.xml");
                                Encoding enc = Encoding.GetEncoding(866);

                                List<StickerInfo> stickerList = (from xe in xdoc.Elements("OrderResponse").Elements("OrderTubes").Elements("OrderTube")
                                                                 select new StickerInfo
                                                                 {
                                                                     BarCodeNumber = xe.Element("LaboratoryNumber").Value,
                                                                     BiomaterialId = xe.Element("BiomaterialId").Value,
                                                                     ZPLText = enc.GetString(Convert.FromBase64String(xe.Element("StickerCodeBase64").Value))
                                                                 }
                                                  ).ToList();
                                List<OrderCoverLetter> letterList = (from xe in xdoc.Elements("OrderResponse").Elements("CoverLetters").Elements("CoverLetter")
                                                                     select new OrderCoverLetter
                                                                     {
                                                                         OrderId = xe.Parent.Parent.Element("OrderId").Value,
                                                                         PdfContent = xe.Element("ContentBase64").Value,
                                                                     }
                                                   ).ToList();

                                Console.WriteLine("Данные для печати получены.");

                                //Создание строк в таблице MSS_BARCODE_NUMBERS
                                /*
                                foreach (var x in stickerList)
                                {
                                    DBControl.InsertBarcodes(qRequest.ID, x, Processing.Ini.OrgId);
                                }

                                //Запись бланка заказа в формате PDF в карту пациента
                                DateTime now = DateTime.Now;
                                string folder = string.Format("{0}-{1}/{2}", now.Year, Convert.ToString(now.Month).PadLeft(2, '0'), qPatient.ID);
                                string pdfDirectory = string.Format("{0}/{1}", Processing.Ini.PdfFilePath, folder);

                                if (!Directory.Exists(pdfDirectory))
                                {
                                    DirectoryInfo dir = Directory.CreateDirectory(pdfDirectory);
                                }

                                int cnt = 1;
                                foreach (var x in letterList)
                                {
                                    Console.WriteLine("Id {0} \n N {1} \n  Code {2} \n ", qRequest.ID, x.OrderId, x.PdfContent);
                                    string fileName = string.Format("{0}_{1}.pdf", x.OrderId, cnt);
                                    string pdfFilePath = string.Format("{0}/{1}", pdfDirectory, fileName);
                                    if (File.Exists(pdfFilePath)) File.Delete(pdfFilePath);
                                    File.WriteAllBytes(pdfFilePath, Convert.FromBase64String(x.PdfContent));
                                    DBControl.AttachPdf(549, qPatient.ID, fileName, now, folder);
                                    cnt++;
                                }
                                //*/
                                //File.WriteAllText(@"D:\data\TestOrders\sticker1905070014.txt", enc.GetString(stickerbytes));
                                //File.WriteAllBytes(@"D:\data\TestOrders\order1905070014.pdf", Convert.FromBase64String(order));
                            }
                        }
                    }
                    //*/
                }
                else { Console.WriteLine("Нет заказов для отправки"); }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Files.WriteLogFile(ex.ToString());
            }
        }

        //Получение данных этикетки из ответного файла
        static public void GetLabelInfo()
        {
            XDocument xdoc = XDocument.Load(@"D:\DATA\Response2205300068.xml");
            string sticker = xdoc.Element("OrderResponse").Element("OrderTubes").Element("OrderTube").Element("StickerCodeBase64").Value;
            var stikerILN = (from st in xdoc.Element("OrderResponse").Element("OrderTubes").Elements("OrderTube").Elements("LaboratoryNumber")
                             select st.Value).ToList();
            foreach (var s in stikerILN) Console.WriteLine(s);

            Console.WriteLine("Данные для печати получены.");

            //System.Drawing.Image stickerImage;
            Encoding enc = Encoding.UTF8;
            byte[] stickerbytes = Convert.FromBase64String(sticker);
            string stickerstring = enc.GetString(stickerbytes);
            File.WriteAllText(@"D:\DATA\toprint.txt", enc.GetString(stickerbytes));
        }
    }
}
