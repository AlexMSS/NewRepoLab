using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
using System.Data.Linq;
using System.Linq;
using System.Xml.Linq;
using System.Net;
using System.IO;
using System.Timers;

namespace LabComm
{
    public class LabControl
    {
        public class Product
        {
            public string Code;
            public string Label;
            public string ExternalCode;
            public int OptionalGroup;
        }

        public class BioMaterial
        {
            public string Code;
            public string Label;
            public int Type;
        }

        public class Container
        {
            public string Code;
            public string Label;
        }

        public class BioContainerMatch
        {
            public string ContainerCode;
            public string BioCode;
            public  bool Obligatory;
        }

        public class ProductBiomat
        {
            public string PrCode;
            public string Option;
            public string BmCode; 
            public bool Obligatory;
            public int Packing;
        }

        public class Optional
        {
            public int OptionalGroup;
            public string Field;
            public string Type;
            public bool Obligatory;
            public string Tag;
        }

        public static System.Timers.Timer DictTimer;

        public static void SetDictTimer()
        {
            DictTimer = new System.Timers.Timer(3600000); //1 
            DictTimer.Elapsed += (obj, e) => GetDictionary(0, e);
            DictTimer.Enabled = false;
        }

        static public void GetDictionary(object obj, ElapsedEventArgs e)
        {
            DateTime dd = DateTime.Now;
            if (dd.Hour == 22 && dd.Minute >= 01 && dd.Minute <= 59)
            {
                Console.WriteLine("Check version {0}", e.SignalTime);
                if (Processing.Ini.DriverCode == "ALab") ALab.GetDict();
            }
        }

        public class ALab
        {
            static public void GetDict()
            {
                try
                {
                    int Ver = GetDictionaryVersion();
                    if (Ver > Processing.Ini.Version)
                    {
                        Processing.Ini.Version = Ver;
                        LoadDictionaries();
                    }
                }
                catch (Exception ex)
                {
                    Files.WriteLogFile("Ошибка при обновлении справочников: " + ex.Message); Console.WriteLine(ex.Message);
                }
            }

            static public int GetDictionaryVersion()
            {
                int Ver = 0;
                XElement messagebody = new XElement("Root");
                Console.WriteLine("Запрос версии справочников от Web"); //Files.WriteLogFile("Запрос версии от Web");

                Message.ALab getversion = new Message.ALab();
                if (getversion.Request("query-dictionaries-version", messagebody, out XDocument responce, out string errmessage))
                {
                    Ver = Convert.ToInt32(responce.Element("Message").Element("Version").Attribute("Version").Value);
                    Console.WriteLine("Текущая версия справочников: {0}, загруженная версия: {1}", Ver, Processing.Ini.Version);
                }

                if (Ver > Processing.Ini.Version)
                {
                    Console.WriteLine("Требуется обновление справочников");
                    List<string> Param = new List<string> { "Version" };
                    List<string> Val = new List<string> { Ver.ToString() };
                    Files.RewriteIniFile(Param, Val);
                }
                else Console.WriteLine("Обновления справочников не требуется");
                return Ver;
            }

            static public void LoadDictionaries()
            {
                XElement messagebody = new XElement("Root");

                Message.ALab getversion = new Message.ALab();
                if (getversion.Request("query-dictionaries", messagebody, out XDocument xdoc, out string errmessage)) Console.WriteLine("Загрузка справочников в базу данных...");

                //Profiles Loading
                Console.WriteLine("Start Profiles Loading....");

                DataContext db = new DataContext(Processing.sqlConnectionStr);

                Table<IdValues> TIdValue = db.GetTable<IdValues>();
                Table<Profiles> TProfiles = db.GetTable<Profiles>();

                var qIdValueP = (from idv in TIdValue
                                 where idv.KeyName == "MSS_PROFILES"
                                 select idv).Single();

                int qId = qIdValueP.LastValue + 1;

                var qDictP = (from xe in xdoc.Element("Message").Elements("Analyses").Elements("Item")
                              select Convert.ToInt32(xe.Attribute("Id").Value))
                            .Except(
                             from pe in TProfiles
                             select pe.Id);

                var insetProfiles = from xe in xdoc.Element("Message").Elements("Analyses").Elements("Item")
                                    join de in qDictP on Convert.ToInt32(xe.Attribute("Id").Value) equals (int?)de
                                    select new Profiles
                                    {
                                        ProfileId = qId++,
                                        OrgId = Processing.Ini.OrgId,
                                        Id = Convert.ToInt32(xe.Attribute("Id").Value),
                                        Code = xe.Attribute("Code").Value,
                                        Label = xe.Attribute("Name").Value
                                            .Substring(0, xe.Attribute("Name").Value.Length < 250 ? xe.Attribute("Name").Value.Length : 250)
                                    };

                TProfiles.InsertAllOnSubmit(insetProfiles);
                qIdValueP.LastValue = qId - 1;
                db.SubmitChanges();

                Console.WriteLine("Profiles have loaded");
                //Biomaterial Loading
                Console.WriteLine("Start Biomaterial Loading....");
                Table<Biomaterial> TBiomaterial = db.GetTable<Biomaterial>();

                var qIdValueB = (from idv in TIdValue
                                 where idv.KeyName == "LAB_BIOTYPE"
                                 select idv).Single();

                qId = qIdValueB.LastValue + 1;

                var qDictB = (from xe in xdoc.Element("Message").Elements("Biomaterials").Elements("Item")
                              select Convert.ToInt32(xe.Attribute("Id").Value))
                             .Except(
                             from pe in TBiomaterial
                             select pe.Id ?? -1);

                var insertBiomaterials = from xe in xdoc.Element("Message").Elements("Biomaterials").Elements("Item")
                                         join de in qDictB on Convert.ToInt32(xe.Attribute("Id").Value) equals (int?)de
                                         select new Biomaterial
                                         {
                                             LabBiotypeId = qId++,
                                             Id = Convert.ToInt32(xe.Attribute("Id").Value),
                                             Code = xe.Attribute("Code").Value,
                                             Label = xe.Attribute("Name").Value,
                                             Localization = "А-Лаб"
                                         };

                TBiomaterial.InsertAllOnSubmit(insertBiomaterials);
                qIdValueB.LastValue = qId - 1;
                db.SubmitChanges();
                Console.WriteLine("Biomaterial have loaded");

                //ProfileBiomaterial Loading

                Console.WriteLine("Start ProfileBiomaterial Loading....");

                Table<ProfilesBiomaterial> TProfileBiomaterial = db.GetTable<ProfilesBiomaterial>();

                var qIdValuePB = (from idv in TIdValue
                                  where idv.KeyName == "MSS_PROFILES_BIOMATERIAL"
                                  select idv).Single();

                qId = qIdValuePB.LastValue + 1;

                var qDictPB = (from xe in xdoc.Elements("Message").Elements("Analyses").Elements("Item").Elements("AnalysisBiomaterials").Elements("Item")
                               select xe.Parent.Parent.Attribute("Id").Value + xe.Attribute("BiomaterialId").Value
                                 ).Except(
                                   from pbe in TProfileBiomaterial
                                   join be in TBiomaterial on pbe.LabBiotypeId equals be.LabBiotypeId
                                   join pe in TProfiles on pbe.ProfileId equals pe.ProfileId
                                   select pe.Id.ToString() + be.Id.ToString());

                var insertPB = from xe in xdoc.Elements("Message").Elements("Analyses").Elements("Item").Elements("AnalysisBiomaterials").Elements("Item")
                               join de in qDictPB on (xe.Parent.Parent.Attribute("Id").Value + xe.Attribute("BiomaterialId").Value) equals de
                               join pe in TProfiles on Convert.ToInt32(xe.Parent.Parent.Attribute("Id").Value) equals pe.Id
                               join be in TBiomaterial on Convert.ToInt32(xe.Attribute("BiomaterialId").Value) equals be.Id
                               select new ProfilesBiomaterial
                               {
                                   ProfilesBiomaterialId = qId++,
                                   LabBiotypeId = be.LabBiotypeId,
                                   ProfileId = pe.ProfileId
                               };

                TProfileBiomaterial.InsertAllOnSubmit(insertPB);
                qIdValuePB.LastValue = qId - 1;
                db.SubmitChanges();
                Console.WriteLine("ProfileBiomaterial have loaded");

                //LabMethods Loading

                Console.WriteLine("Start LabMethods loading....");

                Table<Tests> TTests = db.GetTable<Tests>();

                var qIdValueM = (from idv in TIdValue
                                 where idv.KeyName == "LAB_METHODS"
                                 select idv).Single();

                qId = qIdValueM.LastValue + 1;

                var qDictM = (from xe in xdoc.Elements("Message").Elements("Tests").Elements("Item")
                              select Convert.ToInt32(xe.Attribute("Id").Value)
                                 ).Except(
                                   from me in TTests
                                   where me.GroupID == Processing.Ini.AnalyzerID
                                   select me.Id);

                var insertTests = from xe in xdoc.Elements("Message").Elements("Tests").Elements("Item")
                                  join de in qDictM on Convert.ToInt32(xe.Attribute("Id").Value) equals de
                                  select new Tests
                                  {
                                      LabMethodId = qId++,
                                      Id = Convert.ToInt32(xe.Attribute("Id").Value),
                                      Code = xe.Attribute("Code").Value,
                                      GroupID = Processing.Ini.AnalyzerID,
                                      Label = xe.Attribute("Name").Value
                                            .Substring(0, xe.Attribute("Name").Value.Length < 250 ? xe.Attribute("Name").Value.Length : 250)
                                  };

                TTests.InsertAllOnSubmit(insertTests);
                qIdValueM.LastValue = qId - 1;
                db.SubmitChanges();
                Console.WriteLine("LabMethods have loaded");

                //Drugs as LabMethods Loading

                Console.WriteLine("Start Drugs as LabMethods loading....");

                Table<LAB_DEVS> TLAB_DEVS = db.GetTable<LAB_DEVS>();

                qIdValueM = (from idv in TIdValue
                             where idv.KeyName == "LAB_METHODS"
                             select idv).Single();

                qId = qIdValueM.LastValue + 1;

                qDictM = (from xe in xdoc.Elements("Message").Elements("Drugs").Elements("Item")
                          select Convert.ToInt32(xe.Attribute("Id").Value)
                                 ).Except(
                                   from me in TTests
                                   join ge in TLAB_DEVS on me.GroupID equals ge.LabDevsId
                                   where ge.Label == "Drugs"
                                   select me.Id);

                insertTests = from xe in xdoc.Elements("Message").Elements("Drugs").Elements("Item")
                              join de in qDictM on Convert.ToInt32(xe.Attribute("Id").Value) equals de
                              select new Tests
                              {
                                  LabMethodId = qId++,
                                  Id = Convert.ToInt32(xe.Attribute("Id").Value),
                                  Code = xe.Attribute("Code").Value,
                                  GroupID = (from ge in TLAB_DEVS where ge.Label == "Drugs" select ge.LabDevsId).Single(),
                                  Label = xe.Attribute("Name").Value
                              };

                TTests.InsertAllOnSubmit(insertTests);
                qIdValueM.LastValue = qId - 1;
                db.SubmitChanges();
                Console.WriteLine("Drugs as LabMethods have loaded");
                //Drugs as Dictionaries Loading
                Console.WriteLine("Start Drugs as Dictionaries loading....");
                int rowsAffected = db.ExecuteCommand(@"declare	@count int, @id_par int, @id_bio int, @id_out int, @id_dct int, @id_rel int
                    declare @Labels table(
				                    NUM int,
				                    LAB_METHODS_ID int,
				                    LABEL varchar(250),
				                    CODE varchar(50)
				                    )
                    insert @Labels
                    select ROW_NUMBER() over(order by lm.LAB_METHODS_ID), lm.LAB_METHODS_ID,  lm.LABEL, lm.CODE
                    from LAB_METHODS lm 
                    join LAB_DEVS d on lm.LAB_DEVS_ID=d.LAB_DEVS_ID
                    left outer join LAB_METHODBIO lmb on lm.LAB_METHODS_ID=lmb.LAB_METHODS_ID
                    where lmb.LAB_METHODBIO_ID is null and d.LABEL='Drugs'

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

                    declare @gr_id int = (select top 1 DS_OUTER_DICTPARAMS_ID from DS_OUTER_DICTPARAMS where NUM ='Drugs' )
                    exec dbo.[up_get_id] 'DS_OUTER_RELATIONS',@count,@id_rel output
                    declare @i int = 1	
                    while (@i<=@count)
                    Begin	
	                    insert into DS_OUTER_RELATIONS  (DS_OUTER_RELATIONS_ID,PARENT_ID,CHILD_ID,REL_LEVEL, IS_PRIMARY)
	                    select l.NUM - 1 + @id_rel, @gr_id, dp.DS_OUTER_DICTPARAMS_ID, 0,1
	                    from @Labels l
	                    join DS_OUTER_PARAMS p on l.CODE=p.CODE
	                    join DS_OUTER_DICTPARAMS dp on p.DS_OUTER_PARAMS_ID = dp.DS_OUTER_PARAMS_ID
	                    where l.NUM = @i and p.DS_OUTER_PARAMS_ID>=@id_out 
	                    set @i=@i+1
                    continue
                    end");
                Console.WriteLine("Drugs as Dictionaries have loaded");
                db.Dispose();

            }
        }

        //Загрузка справочников Ариадна
        public class Ariadna
        {
            //Загрузка и обновление справочников
            public static void LoadDictinaries()
            {
                XDocument xdoc;
                DataContext db = new DataContext(Processing.sqlConnectionStr);

                //Справочник услуг с типами биоматериалов
                string filepath;
                if (Directory.Exists(Processing.Ini.DictFilePath))
                    filepath = Directory.GetFiles(Processing.Ini.DictFilePath, "*.xml").FirstOrDefault();
                else
                    filepath = null;
                Console.WriteLine();
                if (filepath != null)
                {
                    xdoc = XDocument.Load(filepath);
                    Console.WriteLine("Загрузка справочника услуг из файла {0}", filepath);

                    //Импорт профилей
                    var profiles = from prof in xdoc.Elements("ServiceList").Elements("Service")
                                   select new Product
                                   {
                                       Code = prof.Element("ID").Value,
                                       Label = prof.Element("Title").Value,
                                       ExternalCode = prof.Element("ExternalID").Value
                                   };
                    
                    foreach (var profile in profiles)
                    {
                        int success = db.ExecuteCommand(@"
                            declare @id int
                            if not exists (select code from MSS_PROFILES where CODE = {0} and FM_ORG_ID = {2} )
                            begin
	                            exec up_get_id 'MSS_PROFILES',1,@id out
                                insert MSS_PROFILES(MSS_PROFILES_ID,  CODE, LABEL , FM_ORG_ID, EXTERNAL_ID)
                                values (@id, {0}, cast({1} as varchar(250)), {2}, {3})
                             end
                            ", profile.Code, profile.Label, Processing.Ini.OrgId, profile.ExternalCode);
                    }


                    //Импорт биоматериалов
                    var biomaterials = (from biomat in xdoc.Elements("ServiceList").Elements("Service").Elements("SpecimenTypes").Elements("SpecimenType")
                                       select new Biomaterial
                                       {
                                           Code = biomat.Element("ID").Value,
                                           Label = biomat.Element("Title").Value,
                                       }).GroupBy(x => new  {x.Code, x.Label }).Select(x => new Biomaterial {Code = x.Key.Code, Label = x.Key.Label });

                    foreach (var bm in biomaterials)
                    {
                        int success = db.ExecuteCommand(@"
                            declare @id int
                            if not exists (select code from LAB_BIOTYPE where CODE = {0})
                            begin
	                            exec up_get_id 'LAB_BIOTYPE', 1, @id out
                                insert LAB_BIOTYPE(LAB_BIOTYPE_ID,  CODE, LABEL)
                                values (@id, {0}, {1})
                             end
                            ", bm.Code, bm.Label);//*/
                    }

                    //Импорт связки профиль-биоматериал
                    var ProfBm = from profbm in xdoc.Elements("ServiceList").Elements("Service").Elements("SpecimenTypes").Elements("SpecimenType")
                                 select new ProductBiomat
                                 {
                                     PrCode = profbm.Parent.Parent.Element("ID").Value,
                                     BmCode = profbm.Element("ID").Value
                                 };
                    

                    foreach (var prbm in ProfBm)
                    {
                        int success = db.ExecuteCommand(@"
                          declare @id int
                          if not exists (select pb.MSS_PROFILES_BIOMATERIAL_ID from MSS_PROFILES_BIOMATERIAL pb
                          join MSS_PROFILES p on pb.MSS_PROFILES_ID = p.MSS_PROFILES_ID
                          join LAB_BIOTYPE b on pb.LAB_BIOTYPE_ID = b.LAB_BIOTYPE_ID
                          where b.CODE = {0} and p.CODE = {1})
                          begin
	                          exec up_get_id 'MSS_PROFILES_BIOMATERIAL',1,@id out
                              insert MSS_PROFILES_BIOMATERIAL (MSS_PROFILES_BIOMATERIAL_ID, MSS_PROFILES_ID, LAB_BIOTYPE_ID)
                              values ( @id 
		                             ,(select MSS_PROFILES_ID from MSS_PROFILES where CODE = {1})
		                             , (select LAB_BIOTYPE_ID from LAB_BIOTYPE where CODE = {0})
                                     )
                          end
                     ", prbm.BmCode, prbm.PrCode);
                    }

                    Console.WriteLine("Профили загружены в базу данных");
                }
                else Console.WriteLine("Нет справочников для загрузки");

                db.Dispose();
            }
        }
        
        //Загрузка справочников КДЛ
        public class KDL
        {
            public static void LoadDitionaries()
            {
                Console.WriteLine("Загрузка справочников....");

                try
                {
                    DataContext db = new DataContext(Processing.sqlConnectionStr);
                    int success;
                    Stream fstream;
                    XDocument xdoc;
                    string dicfolder = @"/%2fLists";
                    string dicname;

                    //Загрузка справочника методик (methodics)
                    ///*
                    dicname = "/methodics.xml";
                    fstream = Files.FTPDownload(dicfolder + dicname, true);
                    xdoc = XDocument.Load(fstream);

                    var methods = from x in xdoc.Elements("root").Elements("methodic")
                                  select new Product
                                  {
                                      Label = x.Element("methodicName").Value,
                                      Code = x.Element("methodicShortName").Value,
                                  };

                    foreach (var x in methods)
                    {
                        success = db.ExecuteCommand(@"
                                declare @id int
                                if not exists (select code from LAB_METHODS 
                                    where CODE = {0} and LAB_DEVS_ID = {2})
                                begin
                                    exec up_get_id 'LAB_METHODS', 1, @id out
                                    insert LAB_METHODS(LAB_METHODS_ID,  CODE, LABEL, LAB_DEVS_ID)
                                    values (@id, {0}, cast({1} as varchar(250)), {2})
                                 end
                            ", x.Code, x.Label, Processing.Ini.AnalyzerID, x.OptionalGroup);
                    }
                    Console.WriteLine("Справочник \"Методик\" загружен");
                    //*/
                    /*
                    //Загрузка справочника типы биометриалов (locuses)
                    dicname = "/locuses.xml";
                    fstream = Files.FTPDownload(dicfolder + dicname);
                    xdoc = XDocument.Load(fstream);
                    var biomat = from x in xdoc.Elements("root").Elements("locus")
                                 select new BioMaterial
                                 {
                                     Label = x.Element("locusName").Value,
                                     Code = x.Element("locusShortName").Value,
                                     Type = Convert.ToInt32(x.Element("locusType").Value)
                                 };

                    foreach (var x in biomat)
                    {
                        success = db.ExecuteCommand(@"
                    declare @id int
                    if not exists (select code from LAB_BIOTYPE 
                    where CODE = {0}  and FM_ORG_ID = {3})
                    begin
	                    exec up_get_id 'LAB_BIOTYPE', 1, @id out
                        insert LAB_BIOTYPE(LAB_BIOTYPE_ID,  CODE, LABEL, TYPE, FM_ORG_ID)
                        values (@id, {0}, {1}, {2}, {3})
                     end
                    ", x.Code, x.Label, x.Type, Processing.Ini.OrgId);
                    }
                    Console.WriteLine("Справочник \"Типы биоматриалов\" загружен");
                    //Загрузка справочника типы контейнеров (containers)
                    dicname = "/containers.xml";
                    fstream = Files.FTPDownload(dicfolder + dicname);
                    xdoc = XDocument.Load(fstream);

                    var contn = from x in xdoc.Elements("root").Elements("container")
                                select new Container
                                {
                                    Label = x.Element("containerName").Value,
                                    Code = x.Element("containerShortName").Value
                                };

                    foreach (var x in contn)
                    {
                        success = db.ExecuteCommand(@"
                    declare @id int
                    if not exists (select code from LAB_CONTTYPE 
                    where CODE = {0} and LABEL = {1})
                    begin
	                    exec up_get_id 'LAB_CONTTYPE', 1, @id out
                        insert LAB_CONTTYPE(LAB_CONTTYPE_ID,  CODE, LABEL)
                        values (@id, {0}, {1})
                     end
                    ", x.Code, x.Label);
                    }
                    Console.WriteLine("Справочник \"Типы контейнеров\" загружен");
                    //Загрузка справочника сопоставления типов биоматериалов с типами контейнеров (Locuses+containers)
                    dicname = "/locus+container.xml";
                    fstream = Files.FTPDownload(dicfolder + dicname);
                    xdoc = XDocument.Load(fstream);

                    var bmContn = from x in xdoc.Elements("root").Elements("locus")
                                  select new BioContainerMatch
                                  {
                                      ContainerCode = x.Element("containerShortName").Value,
                                      BioCode = x.Element("locusShortName").Value,
                                      Obligatory = x.Element("obligatory").Value == "1" ? true : false
                                  };

                    foreach (var x in bmContn)
                    {
                        success = db.ExecuteCommand(@"
                    declare @id int
                    if not exists (
                                select MSS_BIOTYPE_CONTTYPE_ID from MSS_BIOTYPE_CONTTYPE bc
                                join LAB_BIOTYPE bt on bc.LAB_BIOTYPE_ID = bt.LAB_BIOTYPE_ID
                                join LAB_CONTTYPE ct on bc.LAB_CONTTYPE_ID = ct.LAB_CONTTYPE_ID
                                where bt.CODE = {0} and ct.CODE = {1} and bt.FM_ORG_ID = {3}
                                )
                    begin
	                    exec up_get_id 'MSS_BIOTYPE_CONTTYPE', 1, @id out
                        insert MSS_BIOTYPE_CONTTYPE(MSS_BIOTYPE_CONTTYPE_ID, LAB_BIOTYPE_ID, LAB_CONTTYPE_ID, OBLIGATORY)
                        values (  @id
		                        ,(select top 1 LAB_BIOTYPE_ID from LAB_BIOTYPE where CODE = {0} and FM_ORG_ID = {3})
                                ,(select top 1 LAB_CONTTYPE_ID from LAB_CONTTYPE where CODE = {1})
                                , {2})
                     end
                    ", x.BioCode, x.ContainerCode, Convert.ToByte(x.Obligatory), Processing.Ini.OrgId);
                    }
                    Console.WriteLine("Справочник \"Сопоставления типов биоматериалов с типами контейнеров\" загружен");

                    //Загрузка справочника тестов (tests)
                    dicname = "/tests.xml";
                    fstream = Files.FTPDownload(dicfolder + dicname);
                    xdoc = XDocument.Load(fstream);

                    var prod = from x in xdoc.Elements("root").Elements("test")
                               select new Product
                               {
                                   Label = x.Element("testName").Value,
                                   Code = x.Element("testShortName").Value,
                                   OptionalGroup = Convert.ToInt32(x.Element("optionalGroup").Value)
                               };

                    foreach (var x in prod)
                    {
                        success = db.ExecuteCommand(@"
                            declare @id int
                            if not exists (select code from MSS_PROFILES 
                                where CODE = {0} and FM_ORG_ID = {2})
                            begin
	                            exec up_get_id 'MSS_PROFILES', 1, @id out
                                insert MSS_PROFILES(MSS_PROFILES_ID,  CODE, LABEL , FM_ORG_ID, OPTIONAL_GROUP)
                                values (@id, {0}, cast({1} as varchar(250)), {2}, {3})
                             end
                        ", x.Code, x.Label, Processing.Ini.OrgId, x.OptionalGroup);
                    }
                    Console.WriteLine("Справочник \"Профилей\" загружен");

                    //Загрузка справочника сопоставления тестов типам биоматериалов (tests+locuses)
                    dicname = "/test+locus.xml";
                    fstream = Files.FTPDownload(dicfolder + dicname);
                    xdoc = XDocument.Load(fstream);

                    var testbm = from x in xdoc.Elements("root").Elements("test")
                                 select new ProductBiomat
                                 {
                                     PrCode = x.Element("testShortName").Value,
                                     BmCode = x.Element("locusShortName").Value,
                                     Obligatory = x.Element("obligatory").Value == "1" ? true : false,
                                     Packing = Convert.ToInt32(x.Element("packing").Value)
                                 };

                    foreach (var x in testbm)
                    {
                        success = db.ExecuteCommand(@"
                          declare @id int, @b_id int, @p_id int

                          set @p_id = (select top 1 MSS_PROFILES_ID from MSS_PROFILES where CODE = {0} and FM_ORG_ID = {2})
                          set @b_id = (select top 1 LAB_BIOTYPE_ID from LAB_BIOTYPE where CODE = {1} and FM_ORG_ID = {2})

                          if not exists (select pb.MSS_PROFILES_BIOMATERIAL_ID from MSS_PROFILES_BIOMATERIAL pb
                          join MSS_PROFILES p on pb.MSS_PROFILES_ID = p.MSS_PROFILES_ID
                          join LAB_BIOTYPE b on pb.LAB_BIOTYPE_ID = b.LAB_BIOTYPE_ID
                          where b.CODE = {1} and p.CODE = {0}  and p.FM_ORG_ID = {2} and b.FM_ORG_ID = {2} )
                                and @p_id is not null and @b_id is not null
                          begin
	                          exec up_get_id 'MSS_PROFILES_BIOMATERIAL', 1, @id out
                              insert MSS_PROFILES_BIOMATERIAL (MSS_PROFILES_BIOMATERIAL_ID, MSS_PROFILES_ID, LAB_BIOTYPE_ID, OBLIGATORY, PACKING)
                              values ( @id, @p_id, @b_id, {3}, {4} )
                          end
                        ", x.PrCode, x.BmCode, Processing.Ini.OrgId, Convert.ToByte(x.Obligatory), x.Packing);
                    }
                    Console.WriteLine("Справочник \"Сопоставления прфилей типам биоматериалов\" загружен");

                    //Загрузка справочника опций (optional)
                    dicname = "/optional.xml";
                    fstream = Files.FTPDownload(dicfolder + dicname);
                    xdoc = XDocument.Load(fstream);

                    var optional = from x in xdoc.Elements("root").Elements("dopInfo")
                                   select new Optional
                                   {
                                       OptionalGroup = Convert.ToInt32(x.Element("optionalGroup").Value),
                                       Field = x.Element("field").Value,
                                       Type = x.Element("type").Value,
                                       Obligatory = x.Element("obligatory").Value == "1" ? true : false,
                                       Tag = x.Element("tag").Value
                                   };

                    foreach (var x in optional)
                    {
                        success = db.ExecuteCommand(@"
                            declare @id int
                            if not exists (select MSS_PROFILE_OPTION_ID from MSS_PROFILE_OPTION 
                                where OPTIONAL_GROUP = {0} and FIELD = {1})
                            begin
	                            exec up_get_id 'MSS_PROFILE_OPTION', 1, @id out
                                insert MSS_PROFILE_OPTION(MSS_PROFILE_OPTION_ID,  OPTIONAL_GROUP, FIELD , DATA_TYPE, OBLIGATORY, TAG)
                                values (@id, {0}, {1}, {2}, {3}, {4})
                             end
                        ", x.OptionalGroup, x.Field, x.Type, Convert.ToByte(x.Obligatory), x.Tag);
                    }
                    Console.WriteLine("Справочник \"Опции профилей\" загружен");
                    //*/
                }
                catch (Exception ex)
                {
                    Files.WriteLogFile("Ошибка при закгрузке справочников: " + ex.ToString()); Console.WriteLine(ex.Message);
                }

            }
        }
     
    }
}
