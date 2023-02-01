using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Data;
using System.Data.Linq.Mapping;
using System.Data.Linq;
using System.Linq;
using System.Xml.Linq;
using System.Globalization;

namespace LabComm
{
    [Table(Name = "FM_DEP")]
    public class Dep
    {
        [Column(IsPrimaryKey = true, IsDbGenerated = true, Name = "FM_DEP_ID")]
        public int DepId;//{ get; set; }
        [Column(Name = "FM_ORG_ID")]
        public int OrgId;//{ get; set; }
        [Column(Name = "MAIN_ORG_ID")]
        public int MainOrgId;//{ get; set; }
    }

    [Table(Name = "MEDDEP")]
    public class MedDep
    {
        [Column(IsPrimaryKey = true, IsDbGenerated = true, Name = "MEDDEP_ID")]
        public int MedDepId;//{ get; set; }
        [Column(Name = "FM_DEP_ID")]
        public int DepId;//{ get; set; }
        [Column(Name = "MEDECINS_ID ")]
        public int MedecinsId;//{ get; set; }
    }

    [Table(Name = "PATIENTS")]
    public class Patient
    {
        [Column(IsPrimaryKey = true, IsDbGenerated = true, Name = "PATIENTS_ID")]
        public int Id;//{ get; set; }
        [Column(Name = "NOM")]
        public string LastName; //{ get; set; }
        [Column(Name = "PRENOM")]
        public string FirstName; //{ get; set; }
        [Column(Name = "PATRONYME")]
        public string MiddleName; //{ get; set; }
        [Column(Name = "POL")]
        public int Pol; //{ get; set; }
        [Column(Name = "NE_LE")]
        public DateTime BirthDate; //{ get; set; }
    }

    [Table(Name = "PATDIREC")]
    public class Patdirec
    {
        [Column(IsPrimaryKey = true, Name = "PATDIREC_ID")]
        public int Id;//{ get; set; }
        [Column(Name = "BIO_CODE", CanBeNull = true)]
        public string BioCode; //{ get; set; }
        [Column(Name = "DATE_BIO", CanBeNull = true)]
        public Nullable<DateTime> DateBio; //{ get; set; }
        [Column(IsPrimaryKey = true, Name = "PATIENTS_ID")]
        public int PatientID; //{ get; set; }
        [Column(Name = "MSS_ORDER_REQUEST_ID", CanBeNull = true)]
        public int? OrderId;//{ get; set; }
        [Column(Name = "PL_EXAM_ID", CanBeNull = true)]
        public int PlExamId;//{ get; set; }
        [Column(Name = "MOTCONSU_ID")]
        public int MotconsuId; //{ get; set; }
        [Column(Name = "MEDECINS_BIO_DEP_ID")]
        public int MedBioDepId; //{ get; set; }
    }

    [Table(Name = "MOTCONSU")]
    public class Motconsu
    {
        [Column(IsPrimaryKey = true, Name = "MOTCONSU_ID")]
        public int Id;//{ get; set; }
        [Column(Name = "MOTCONSU_EV_ID", CanBeNull = true)]
        public int? EventId;//{ get; set; }
        [Column(Name = "Date_Consultation")]
        public DateTime Date_Consultation { get; set; }
    }

    [Table(Name = "PL_EXAM")]
    public class PlExam
    {
        [Column(IsPrimaryKey = true, Name = "PL_EXAM_ID")]
        public int Id;//{ get; set; }
        [Column(Name = "NAME", CanBeNull = true)]
        public string Name;
        [Column(IsPrimaryKey = true, Name = "MODELS_ID")]
        public int ModelsId; //{ get; set; }
    }

    [Table(Name = "DIR_SERV")]
    public class DirServ
    {
        [Column(IsPrimaryKey = true, Name = "DIR_SERV_ID")]
        public int DirServId;//{ get; set; }
        [Column(Name = "PATDIREC_ID")]
        public int PatdirId; //{ get; set; }
        [Column(Name = "FM_SERV_ID")]
        public int ServId; //{ get; set; }
        [Column(Name = "FM_ORG_ID", CanBeNull = true)]
        public int? OrgId; //{ get; set; }
    }

    [Table(Name = "MSS_ORDER_SERV")]
    public class ServofOrder
    {
        [Column(IsPrimaryKey = true, Name = "MSS_ORDER_SERV_ID")]
        public int ServofOrderId;//{ get; set; }
        [Column(Name = "FM_ORG_ID")]
        public int OrgId; //{ get; set; }
        [Column(Name = "FM_SERV_ID")]
        public int ServId; //{ get; set; }
        [Column(Name = "STATE")]
        public int State; //{ get; set; }
    }

    [Table(Name = "MSS_BARCODE_NUMBERS")]
    public class BarcodeNum
    {
        [Column(IsPrimaryKey = true, Name = "MSS_BARCODE_NUMBERS_ID")]
        public int BarcodeId;//{ get; set; }
        [Column(Name = "FM_ORG_ID")]
        public int OrgId; //{ get; set; }
        [Column(Name = "LAB_BIOTYPE_ID")]
        public int BiotypeId; //{ get; set; }
        [Column(Name = "PATDIREC_ID")]
        public int PatdirId; //{ get; set; }
        [Column(Name = "BARCODE_NUMBER")]
        public string BarcodeNumber; //{ get; set; }
    }

    [Table(Name = "MSS_ORDER_REQUEST")]
    public class OrderRequest
    {
        [Column(IsPrimaryKey = true, Name = "MSS_ORDER_REQUEST_ID")]
        public int OrderId;//{ get; set; }
        [Column(Name = "FM_ORG_ID")]
        public int OrgId; //{ get; set; }
        [Column(Name = "PATIENTS_ID")]
        public int PatientId; //{ get; set; }
        [Column(Name = "DATE_ORDER")]
        public DateTime DateOrder; //{ get; set; }
        [Column(Name = "ORDER_NUMBER")]
        public int OrderNumber; //{ get; set; }
        [Column(Name = "STATE")]
        public int State; //{ get; set; }
        [Column(Name = "RequestCode")]
        public string Code; //{ get; set; }
    }

    [Table(Name = "MSS_PROFILE_SERV")]
    public class ProfileServ
    {
        [Column(IsPrimaryKey = true, Name = "MSS_PROFILE_SERV_ID")]
        public int DirServId;//{ get; set; }
        [Column(Name = "FM_SERV_ID")]
        public int ServId; //{ get; set; }
        [Column(Name = "MSS_PROFILES_ID")]
        public int ProfileId; //{ get; set; }
    }

    [Table(Name = "MSS_PROFILES")]
    public class Profiles
    {
        [Column(IsPrimaryKey = true , Name = "MSS_PROFILES_ID", CanBeNull = true)]
        public int ProfileId { get; set; }
        [Column(Name = "FM_ORG_ID", CanBeNull = true)]
        public int OrgId { get; set; }
        [Column(IsPrimaryKey = true, Name = "EXTERNAL_ID", CanBeNull = true)]
        public int Id { get; set; }
        [Column(Name = "CODE", CanBeNull = true)]
        public string Code { get; set; }
        [Column(Name = "LABEL", CanBeNull = true)]
        public string Label { get; set; }
    }

    [Table(Name = "ID_VALUES")]
    public class IdValues
    {
        [Column(IsPrimaryKey = true, Name = "KEY_NAME")]
        public string KeyName { get; set; }
        [Column(Name = "LAST_VALUE")]
        public int LastValue { get; set; }
    }

    [Table(Name = "LAB_BIOTYPE")]
    public class Biomaterial
    {
        [Column(IsPrimaryKey = true, Name = "LAB_BIOTYPE_ID", CanBeNull = true)]
        public int LabBiotypeId { get; set; }
        [Column(IsPrimaryKey = true, Name = "MSS_BioMaterial_ID", CanBeNull = true)]
        public int? Id { get; set; }
        [Column(Name = "CODE", CanBeNull = true)]
        public string Code { get; set; }
        [Column(Name = "LOCALIZATIONS", CanBeNull = true)]
        public string Localization { get; set; }
        [Column(Name = "LABEL", CanBeNull = true)]
        public string Label { get; set; }

    }

    [Table(Name = "LAB_METHODS")]
    public class Tests
    {
        [Column(IsPrimaryKey = true, Name = "LAB_METHODS_ID", CanBeNull = true)]
        public int? LabMethodId { get; set; }
        [Column(IsPrimaryKey = true, Name = "NUM", CanBeNull = true)]
        public int Id { get; set; }
        [Column(Name = "CODE", CanBeNull = true)]
        public string Code { get; set; }
        [Column(Name = "LAB_DEVS_ID", CanBeNull = true)]
        public int GroupID { get; set; }
        [Column(Name = "LABEL", CanBeNull = true)]
        public string Label { get; set; }
    }

    [Table(Name = "LAB_DEVS")]
    public class LAB_DEVS
    {
        [Column(IsPrimaryKey = true, Name = "LAB_DEVS_ID")]
        public int LabDevsId { get; set; }
        [Column(Name = "LABEL", CanBeNull = true)]
        public string Label { get; set; }
    }

    [Table(Name = "FM_ORG")]
    public class FM_ORG
    {
        [Column(IsPrimaryKey = true, Name = "FM_ORG_ID", CanBeNull = true)]
        public int? OrgId { get; set; }
        [Column(Name = "ORG_TYPE")]
        public string Type { get; set; }
        [Column(Name = "Label")]
        public string Label { get; set; }
    }

    [Table(Name = "MSS_EXTERNAL_ORG_CODES")]
    public class EXT_ORG_CODE
    {
        [Column(IsPrimaryKey = true, Name = "EXTERNAL_ORG_ID", CanBeNull = true)]
        public int? ExtOrgId { get; set; }
        [Column(Name = "CODE")]
        public string Code { get; set; }
        [Column(Name = "FM_ORG_ID")]
        public int OrgId { get; set; }
    }

    [Table(Name = "MSS_PROFILES_BIOMATERIAL")]
    public class ProfilesBiomaterial
    {
        [Column(IsPrimaryKey = true, Name = "MSS_PROFILES_BIOMATERIAL_ID", CanBeNull = true)]
        public int ProfilesBiomaterialId { get; set; }
        [Column(Name = "LAB_BIOTYPE_ID", CanBeNull = true)]
        public int LabBiotypeId { get; set; }
        [Column(Name = "MSS_PROFILES_ID", CanBeNull = true)]
        public int ProfileId { get; set; }
        [Column(Name = "Flag", CanBeNull = true)]
        public int Flag { get; set; }
    }
    // создание таблицы лога
    [Table(Name = "MSS_LOAD_DICT_LOGS")]
    public class ProfilesBiomateriallog
    {
        [Column(IsPrimaryKey = true, Name = "MSS_PROFILES_BIOMATERIAL_ID", CanBeNull = true)]
        public int ProfilesBiomaterialId { get; set; }
        [Column(Name = "LAB_BIOTYPE_ID", CanBeNull = true)]
        public int LabBiotypeId { get; set; }
        [Column(Name = "MSS_PROFILES_ID", CanBeNull = true)]
        public int ProfileId { get; set; }
        [Column(Name = "Change", CanBeNull = true)]
        public int Change { get; set; }
    }
    
    [Table(Name = "IMPDATA")]
    public class Impdata
    {
        [Column(IsPrimaryKey = true, Name = "IMPDATA_ID", CanBeNull = true)]
        public int ImpdataId { get; set; }
        [Column(Name = "ImpDeleted", CanBeNull = true)]
        public int ImpDeleted { get; set; }
        [Column(Name = "Nom", CanBeNull = true)]
        public string Nom { get; set; }
        [Column(Name = "Prenom", CanBeNull = true)]
        public string Prenom { get; set; }
        [Column(Name = "Date_Consultation", CanBeNull = true)]
        public DateTime Date_Consultation { get; set; }
        [Column(Name = "Date_Naissance", CanBeNull = true)]
        public DateTime Date_Naissance { get; set; }
        [Column(Name = "Mesure", CanBeNull = true)]
        public string Mesure { get; set; }
        [Column(Name = "KEYCODE", CanBeNull = true)]
        public string KEYCODE { get; set; }
        [Column(Name = "PATIENTS_ID", CanBeNull = true)]
        public int? PATIENTS_ID { get; set; }
        [Column(Name = "STATE", CanBeNull = true)]
        public int STATE { get; set; }
        [Column(Name = "FILIAL_CODE", CanBeNull = true)]
        public string FilialCode { get; set; }
    }

   [Table(Name = "DS_RESTESTS")]
   public class Restests
   {
        [Column(IsPrimaryKey = true, Name = "DS_RESTESTS_ID", CanBeNull = true)]
        public int RestestsId { get; set; }
        [Column(Name = "IMPDATA_ID", CanBeNull = true)]
        public int ImpdataId { get; set; }
        [Column(Name = "VAL", CanBeNull = true)]
        public string VAL { get; set; }
        [Column(Name = "RES_DATE", CanBeNull = true)]
        public DateTime RES_DATE { get; set; }
        [Column(Name = "UNIT", CanBeNull = true)]
        public string UNIT { get; set; }
        [Column(Name = "STATE", CanBeNull = true)]
        public int STATE { get; set; }
        [Column(Name = "MOTCONSU_ID", CanBeNull = true)]
        public int? MOTCONSU_ID { get; set; }
        [Column(Name = "COMMENT", CanBeNull = true)]
        public string COMMENT { get; set; }
        [Column(Name = "LAB_METHODS_ID", CanBeNull = true)]
        public int LAB_METHODS_ID { get; set; }
        [Column(Name = "RES_TYPE", CanBeNull = true)]
        public string RES_TYPE { get; set; }
        [Column(Name = "M_VAL", CanBeNull = true)]
        public string M_VAL { get; set; }
        [Column(Name = "ITEM", CanBeNull = true)]
        public int? ITEM { get; set; }
        [Column(Name = "NORM_MIN_REC", CanBeNull = true)]
        public decimal? NORM_MIN_REC { get; set; }
        [Column(Name = "NORM_MAX_REC", CanBeNull = true)]
        public decimal? NORM_MAX_REC { get; set; }
        [Column(Name = "NORM_TEXT_REC", CanBeNull = true)]
        public string NORM_TEXT_REC { get; set; }
        [Column(Name = "NORM_STATE", CanBeNull = true)]
        public int NORM_STATE { get; set; }
        [Column(Name = "USE_NORMEXPR", CanBeNull = true)]
        public bool USE_NORMEXPR { get; set; }
        [Column(Name = "RESCODE", CanBeNull = true)]
        public string Rescode { get; set; }
    }
    
    public class DepOrg
    {
        public int DepId;
        public int MedDepId;
        public int MedecinsId;
    }

    public class Biomat
    {
        public string Label;
        public string Code;
    }

    public class Profile
    {
        public string Code;
        public string Label;
        public string ExternalId = string.Empty;
        public string ShortLabel = string.Empty;
        public string GroupName;
        public string SubgroupName;
    }

    public class ProfileBiomat
    {
        public string ProfileCode;
        public string BiomatCode;
        public string OptionSet = string.Empty;
        public string ProfileExtId;
        public string Flag = string.Empty;
    }

    public class ProfileBiomatlog
    {
        public string ProfileCode;
        public string BiomatCode;
        public string OptionSet = string.Empty;
        public string ProfileExtId;
        public string Change;
    }

    public class Test
    {
        public string Label;
        public string Code;
    } 

    public class FilialInfo
    {
        public string Token;
        public string ClientId;
    }

    public class DBControl
    {
        static public SqlDataReader dataReader = null;

        static private string UserID;

        static public string GetSqlConnectionString()
        {
            List<string> ConnectionParameters = new List<string> { "Data Source", "Initial Catalog", "User ID", "Password" };
            List<string> ConnectionProperties = null;
            Files.ReadIniFile(ConnectionParameters, out ConnectionProperties);
            SqlConnectionStringBuilder sqlConnectionSB = new SqlConnectionStringBuilder();

            sqlConnectionSB.DataSource = ConnectionProperties[0];
            sqlConnectionSB.InitialCatalog = ConnectionProperties[1];
            sqlConnectionSB.UserID = ConnectionProperties[2];
            sqlConnectionSB.Password = ConnectionProperties[3];

            UserID = ConnectionProperties[2];

            sqlConnectionSB.ConnectTimeout = 2;

            return sqlConnectionSB.ToString();
        }

        static public string GetSID()
        {
            string str = "";
            try
            {
                DataContext db = new DataContext(Processing.sqlConnectionStr);
                str = db.ExecuteQuery<string>(@"
                    select cast(( master.dbo.fn_varbintohexstr(sid))as nvarchar(100)) sid from master.sys.sql_logins 
                    where name = {0}"
                    , UserID).Single();
            }
            catch (Exception ex)
            {
                Files.WriteLogFile(ex.ToString());
                System.Threading.Thread.Sleep(5000);
                Environment.Exit(1);
            }
            return str;
        }

        public class SIDData
        {
            public string SID;
            public DateTime CreateDate;
        }

        static public string GetSIDPlus()
        {
            SIDData sid = new SIDData();
            Console.WriteLine("Подключение к базе данных...");
            try
            {
                DataContext db = new DataContext(Processing.sqlConnectionStr);
                sid = db.ExecuteQuery<SIDData>(@"
                    select cast(( master.dbo.fn_varbintohexstr(sid))as nvarchar(100)) SID, create_date CreateDate from master.sys.sql_logins 
                    where name = {0}"
                    , UserID).Single();
                db.Dispose();
            }
            catch (Exception ex)
            {
                Files.WriteLogFile(ex.ToString());
                System.Threading.Thread.Sleep(6000);
                Environment.Exit(1);
            }
            Console.WriteLine("Подключение к базе данных выполнено");
            return sid.SID + sid.CreateDate.ToString("yyMMddHHmmssf");

        }

        static public bool GetImpdataID(DataContext db, string KeyCode, out int ImpdataID)
        {
            ImpdataID = db.ExecuteQuery<int>(@"
                    declare @imp_id int
                    set @imp_id = (select top 1 impdata_id 
                    from IMPDATA 
                    where DATEADD(day, 21, KRN_CREATE_DATE) > GETDATE()
                    and KEYCODE = {0}
                    order by KRN_CREATE_DATE desc) 

                    select ISNULL(@imp_id,0)
                    ", KeyCode).Single();

            if (ImpdataID == 0)
            {
                ImpdataID = db.ExecuteQuery<int>(@"
                        declare @imp_id int
                        exec up_get_id 'IMPDATA',1, @imp_id out
                        select @imp_id
                        ", KeyCode).Single();
                return true;
            }
            else
            {
                Files.WriteLogFile(string.Format("Повторный файл {0}", KeyCode));
                Console.WriteLine("Повторный файл {0}", KeyCode);
                return false;
            }

        }

        //Вставка в ImpData
        static public void InsertImpdata(DataContext db, Processing.ImpData impData )
        {
            int ok = db.ExecuteCommand(@"
                insert IMPDATA (Impdata_ID, Date_Consultation, KEYCODE, Nom, Prenom, Date_Naissance, Mesure, QC_INFO, KRN_CREATE_USER_ID, MSG_TYPE, MSS_LAB_CODE)
	            values ({0}, {1}, {2}, {3}, {4}, {5}, {6}, 0, 1,'D', {6})
                ", impData.ImpdataId, impData.Date_Consultation, impData.KEYCODE, impData.Nom, impData.Prenom, impData.Date_Naissance
                        , impData.Mesure);
        }

        //Вставка в DS_RESTESTS
        static public void InsertRestestData(DataContext db, List<Processing.RestestData> qRestest, int ImpdataId, DateTime resdate)
        {
            foreach (var el in qRestest)
            {
                int state = 3;
                if (el.VAL != "" && el.VAL != null)
                {
                    state = db.ExecuteQuery<int>(@"
	                    declare @res_id int, @labm_id int, @state int
	                    
                        declare @lmt table (id int)
                        insert @lmt (id)
                        select LAB_METHODS_ID  from LAB_METHODS 
                        where LAB_DEVS_ID = {5}  and Code = {1}
 
                        declare @cnt int 
                        set @cnt = (select COUNT (*) from @lmt)
                        set @labm_id = case when @cnt = 0 then 0 
		                        else (select MIN(id) from @lmt) end

                        if @labm_id = 0
                        begin
                           exec up_get_id 'LAB_METHODS', 1, @labm_id out
                           insert into LAB_METHODS (LAB_METHODS_ID, LABEL, CODE, UNIT, LAB_DEVS_ID, DATATYPE)
                           values (@labm_id, {7}, {1}, {8}, {5}, {10})
                        end

                        declare @val varchar(max)
                        select @val = case when m_val is null then val else cast(m_val as varchar(max)) end 
                        from DS_RESTESTS 
                        where IMPDATA_ID = {0}
                            and LAB_METHODS_ID = @labm_id
                            and isnull(ITEM, 0) = {11}

                        set @state = 0
                        if (@val is null)
                        begin
	                        set @state = 2
                            exec up_get_id 'ds_restests', 1, @res_id out
	                        insert DS_RESTESTS (DS_RESTESTS_ID, Impdata_ID, VAL, M_VAL, RES_DATE, LAB_METHODS_ID, STATE,
	                        KRN_CREATE_USER_ID, RES_TYPE, NORM_TEXT_REC, UNIT, COMMENT, ITEM, RESCODE)
	                        values (@res_id, {0}, case when len({2}) > 255 then null else {2} end
                                                , case when len({2}) > 255 then {2} else null end
                                , {3}, @labm_id, {4}, 1 ,'D', {6}, {8}, {9}
                                    , case when {11} = 0 then null else {11} end, {12})
                        end
                        else 
                        if (@val <> {2})
                            begin
                                set @state = 1
                                if (len({2})<255)
                                update DS_RESTESTS set VAL = {2}
                                where IMPDATA_ID = {0}
                                    and LAB_METHODS_ID = @labm_id
                                    and isnull(ITEM, 0) = {11}
                                else 
                                update DS_RESTESTS set M_VAL = {2}
                                where IMPDATA_ID = {0}
                                    and LAB_METHODS_ID = @labm_id
                                    and isnull(ITEM, 0) = {11}
                            end 
                        select @state
                    ", ImpdataId, el.Rescode, el.VAL, resdate, el.State, Processing.Ini.AnalyzerID, el.NORM_TEXT_REC ?? string.Empty
                            , el.ResName ?? string.Empty, el.UNIT ?? string.Empty, el.Comments ?? string.Empty, el.DataType, el.Item, el.ResPage ?? string.Empty
                         ).Single();
                }
                if (state == 2)  Console.WriteLine("Новые результаты\n{0}\t{1}\t{2}", el.Rescode, el.VAL, el.UNIT);
                if (state == 1) Console.WriteLine("Результаты обновлены\n{0}\t{1}\t{2}", el.Rescode, el.VAL, el.UNIT);
                if (state == 0) Console.WriteLine("Дублирование результатов");
                if (state == 3) Console.WriteLine("Пусто");
            }
        }

        //Вставка в DS_RESTESTS
        static public void InsertRestestDataIDM(DataContext db, List<Processing.RestestData> qRestest, int ImpdataId, DateTime resdate)
        {
            foreach (var el in qRestest)
            {
                int state = db.ExecuteQuery<int>(@"
	                    declare @res_id int, @state int, @imp_id int, @labm_id int, @rval varchar(1000), @resdate datetime, @rstate int
                                , @norms varchar(1000), @unit varchar(30), @comments varchar(1000), @item int, @respage varchar(255)
	                     
							set @imp_id = {0}
							set @labm_id = {1}
							set @rval = {2}
							set @resdate = {3}
							set @rstate = {4}
							set @norms = {5}
							set @unit = {6}
							set @comments = {7}
                            set @item = {8}
							set @respage = {9}

                        declare @val varchar(max)
                        select @val = case when m_val is null then val else cast(m_val as varchar(max)) end 
                        from DS_RESTESTS 
                        where IMPDATA_ID = @imp_id
                            and LAB_METHODS_ID = @labm_id
                            and isnull(ITEM, 0) = @item

                        set @state = 0
                        if (@val is null)
                        begin
	                        set @state = 2
                            exec up_get_id 'ds_restests', 1, @res_id out
	                        insert DS_RESTESTS (DS_RESTESTS_ID, Impdata_ID, VAL, M_VAL, RES_DATE, LAB_METHODS_ID, STATE,
	                        KRN_CREATE_USER_ID, RES_TYPE, NORM_TEXT_REC, UNIT, COMMENT, ITEM, RESCODE)
	                        values (@res_id, @imp_id, case when len(@rval) > 255 then null else @rval end
                                                , case when len(@rval) > 255 then @rval else null end
                                , @resdate, @labm_id, @rstate, 1 ,'D', @norms, @unit, @comments
                                    , case when @item = 0 then null else @item end, @respage)
                        end
                        else 
                        if (@val <> @rval)
                            begin
                                set @state = 1
                                if (len(@rval)<255)
                                update DS_RESTESTS set VAL = @rval
                                where IMPDATA_ID = @imp_id
                                    and LAB_METHODS_ID = @labm_id
                                    and isnull(ITEM, 0) = @item
                                else 
                                update DS_RESTESTS set M_VAL = @rval
                                where IMPDATA_ID = @imp_id
                                    and LAB_METHODS_ID = @labm_id
                                    and isnull(ITEM, 0) = @item
                            end 
                        select @state
                    ", ImpdataId, el.MethodId, el.VAL, resdate, el.State, el.NORM_TEXT_REC ?? string.Empty
                        , el.UNIT ?? string.Empty, el.Comments ?? string.Empty, el.Item, el.ResPage ?? string.Empty
                     ).Single();
                if (state == 2) Console.WriteLine("Новые результаты\n{0}\t{1}\t{2}", el.Rescode, el.VAL, el.UNIT);
                if (state == 1) Console.WriteLine("Результаты обновлены\n{0}\t{1}\t{2}", el.Rescode, el.VAL, el.UNIT);
                if (state == 0) Console.WriteLine("Дублирование результатов");
            }
        }

        //Вставка в DS_RESTESTS 
        static public void InsertRestestItem(DataContext db, Processing.RestestData restestData, int ImpdataId, DateTime resdate)
        {
            int state = db.ExecuteQuery<int>(@"
	                    declare @res_id int, @state int, @labm_id int
	                    
                        set @labm_id = (select TOP 1 LAB_METHODS_ID  
                                        from LAB_METHODS 
                                        where LAB_DEVS_ID = {5}  
                                            and CODE = {1}
                                        )
                        set @state = 0
                        if ({2} <> '')
                        begin
                            declare @val varchar(250)
                            select @val = VAL 
                            from DS_RESTESTS 
                            where IMPDATA_ID = {0}
                            and LAB_METHODS_ID = @labm_id

                            if (@val is null)
                            begin
	                            set @state = 2
                                exec up_get_id 'ds_restests', 1, @res_id out
	                            insert DS_RESTESTS (DS_RESTESTS_ID, Impdata_ID, VAL, RES_DATE, LAB_METHODS_ID, STATE,
	                            KRN_CREATE_USER_ID, RES_TYPE)
	                            values (@res_id, {0}, {2}, {3}, @labm_id, {4}, 1 ,'D')
                            end
                            else 
                            if (@val <> {2})
                            begin
                                set @state = 1
                                update DS_RESTESTS set VAL = {2}
                                where IMPDATA_ID = {0}
                                    and LAB_METHODS_ID = @labm_id
                                end 
                        end
                       
                        if ({6} <> '')
                        begin
                            declare @m_val varchar(max)
                            select @m_val = M_VAL 
                            from DS_RESTESTS 
                            where IMPDATA_ID = {0}
                            and LAB_METHODS_ID = @labm_id

                            if (@m_val is null)
                            begin
	                            set @state = 2
                                exec up_get_id 'ds_restests', 1, @res_id out
	                            insert DS_RESTESTS (DS_RESTESTS_ID, Impdata_ID, RES_DATE, LAB_METHODS_ID, STATE,
	                            KRN_CREATE_USER_ID, RES_TYPE, M_VAL)
	                            values (@res_id, {0}, {3}, @labm_id, {4}, 1 ,'D', {6})
                            end
                            else 
                            if (@m_val <> {6})
                            begin
                                set @state = 1
                                update DS_RESTESTS set M_VAL = {6}
                                where IMPDATA_ID = {0}
                                    and LAB_METHODS_ID = @labm_id
                                end 
                        end
                        select @state
                    ", ImpdataId, restestData.Rescode, restestData.VAL ?? string.Empty, resdate, 0, Processing.Ini.AnalyzerID, restestData.M_VAL ?? string.Empty
                    ).Single();
            if (state == 2) Console.WriteLine("Новые результаты\n{0}\t{1}", restestData.Rescode, restestData.VAL);
            if (state == 1) Console.WriteLine("Результаты обновлены\n{0}\t{1}", restestData.Rescode, restestData.VAL);
            if (state == 0) Console.WriteLine("Дублирование результатов");
        }

        //Обновление Impdata
        static public void UpdateImpdata(DataContext db, int ImpdataId, int PatId)
        {
            int rowsAffected = db.ExecuteCommand(@"
	                update IMPDATA set PATIENTS_ID={1}, STATE=1
	                where IMPDATA_ID={0}
                ", ImpdataId, PatId);
        }

        //Определение ID отделения  по коду в XML ответе
        static public void GetDepId(string DepCode, int MedId, out int DepId, out int MedDepId)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            DepOrg dep = db.ExecuteQuery<DepOrg>(@"
                    select ext.FM_DEP_ID DepId, mdd.MEDDEP_ID MedDepId
                    from MSS_EXTERNAL_ORG_CODES ext
                    join MEDDEP mdd on ext.FM_DEP_ID = mdd.FM_DEP_ID
                    where EXTERNAL_ORG_ID = {0}
	                    and DEP_CODE = {1}
	                    and mdd.MEDECINS_ID = {2}
                    ", Processing.Ini.OrgId, DepCode, MedId).SingleOrDefault();

            if (dep != null)
            {
                DepId = dep.DepId;
                MedDepId = dep.MedDepId;
            }
            else
            {
                DepId = Processing.DepId;
                MedDepId = Processing.MedDepId;
                Console.WriteLine("Отделение пользователя не определено!");
            }
        }

        static public bool GetDepId(out int MedicinsId, out int DepId, out int MedDepId)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            bool result = true;
            DepOrg dep = db.ExecuteQuery<DepOrg>(@"select dep.FM_DEP_ID DepId, mdd.MEDDEP_ID MedDepId, mdd.MEDECINS_ID MedecinsId
                     from FM_DEP dep
					 join MEDDEP mdd on dep.FM_DEP_ID = mdd.FM_DEP_ID
					 where dep.FM_ORG_ID = {0}
                    ", Processing.Ini.OrgId).SingleOrDefault();
            if (dep == null)
            {
                result = false;
                DepId = 0;
                MedDepId = 0;
                MedicinsId = 0;
            }
            else
            {
                DepId = dep.DepId;
                MedDepId = dep.MedDepId;
                MedicinsId = dep.MedecinsId;
            }
            return result;
        }
        //Отправка в историю болезни (создание записей в таблице MOTCONSU)
        static public void InsertToMotconsu(int ImpdataId, int p_id, DateTime resdate, bool NewImport)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            if (NewImport)
            {
                if (p_id != 0)
                {
                    int rowsAffected = db.ExecuteCommand(@"
                           	declare @m_id int, @ev_id int, @DATEtime datetime 
                            set @DATEtime = Getdate()
                            
		                    set	@m_id=(select top 1 m.MOTCONSU_ID from MOTCONSU m
		                    where PATIENTS_ID={1} 
		                    and convert (varchar(6), m.DATE_CONSULTATION, 12 ) = convert (varchar(6), {2}, 12 )
		                    and m.MODELS_ID={3} and m.MEDECINS_ID={4} )
		
		                   if @m_id is null 
                              or exists 
		                        (select drImp.LAB_METHODS_ID 
		                        from IMPDATA imp
		                        join DS_RESTESTS drImp on imp.Impdata_ID = drImp.IMPDATA_ID
		                        where imp.Impdata_ID = {0} 
				                        and drImp.LAB_METHODS_ID = 
			                        ANY (
			                        select dr.LAB_METHODS_ID 
			                        from MOTCONSU m
			                        join DS_RESTESTS dr on m.MOTCONSU_ID = dr.MOTCONSU_ID
			                        where m.MOTCONSU_ID = @m_id
			                        )
		                        )
		                     begin
			                    set @ev_id=(SELECT TOP 1 MOTCONSU.MOTCONSU_EV_ID
			                    FROM PATDIREC PATDIREC 
			                    JOIN MOTCONSU MOTCONSU ON MOTCONSU.MOTCONSU_ID = PATDIREC.MOTCONSU_ID 
			                    JOIN PL_EXAM PL_EXAM ON PL_EXAM.PL_EXAM_ID = PATDIREC.PL_EXAM_ID
			                    WHERE PATDIREC.PATIENTS_ID= {1}   
				                    and datediff(day, PATDIREC.DATE_BIO,  @DATEtime)<= PL_EXAM.VAL_PERIOD 
				                    and PL_EXAM.PL_EX_GR_ID=33
                                order by PATDIREC.PATDIREC_ID)

                                exec up_get_id 'MOTCONSU',1,@m_id out
			                    INSERT into MOTCONSU (MOTCONSU_ID,    PATIENTS_ID,  MEDECINS_ID, FM_DEP_ID,MEDECINS_CREATE_ID, 
			                    MODELS_ID, DATE_CONSULTATION,  CREATE_DATE_TIME, REC_STATUS, MODIFY_DATE_TIME, MEDECINS_MODIFY_ID, KRN_MODIFY_USER_ID, 
			                    MOTCONSU_EV_ID, REC_NAME, MEDDEP_ID) 
			                    values (@m_id, {1} , {4}, {5}, {4}, {3}, {2}, @DATEtime,'W', @DATEtime, {4}, {4}, 
			                    @ev_id, {7}, {6})
		                    end

		                 --Обновление IMPDATA
		                 update IMPDATA set STATE=1
		                 where Impdata_ID={0} 
		                 --Обновление DS_RESTESTS
		                 update DS_RESTESTS set MOTCONSU_ID = @m_id
		                 where Impdata_ID={0}"
                     , ImpdataId, p_id, resdate, Processing.Ini.ModelId, Processing.MedicinsId, Processing.DepId, Processing.MedDepId, Processing.Mesure);
                }
            }
            else
            {
                int rowsAffected = db.ExecuteCommand(@"
                        update DS_RESTESTS set MOTCONSU_ID = 
                            (select top 1 dr.MOTCONSU_ID
                            from IMPDATA imp
                            join DS_RESTESTS dr on imp.Impdata_ID = dr.IMPDATA_ID
                            where imp.Impdata_ID = {0})
                        where MOTCONSU_ID is null
                            and Impdata_ID = {0}"
                     , ImpdataId);
            }

            db.Dispose();
        }

        static private int InsertToEMC(int ImpdataId, int p_id, DateTime resdate, bool NewImport)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int motconsuId = 0;

            if (NewImport)
            {
                if (p_id != 0)
                {
                    motconsuId = db.ExecuteQuery<int>(@"
                           	declare @m_id int, @ev_id int, @DATEtime datetime 
                            set @DATEtime = Getdate()
                            
		                    set	@m_id = null
                            /*(select top 1 m.MOTCONSU_ID from MOTCONSU m
		                    where PATIENTS_ID={1} 
		                    and convert (varchar(6), m.DATE_CONSULTATION, 12 ) = convert (varchar(6), {2}, 12 )
		                    and m.MODELS_ID={3} and m.MEDECINS_ID={4} and m.MEDDEP_ID = {6}  
                            )*/
		
		                   if @m_id is null 
                              or exists 
		                        (select drImp.LAB_METHODS_ID 
		                        from IMPDATA imp
		                        join DS_RESTESTS drImp on imp.Impdata_ID = drImp.IMPDATA_ID
		                        where imp.Impdata_ID = {0} 
				                        and drImp.LAB_METHODS_ID = 
			                        ANY (
			                        select dr.LAB_METHODS_ID 
			                        from MOTCONSU m
			                        join DS_RESTESTS dr on m.MOTCONSU_ID = dr.MOTCONSU_ID
			                        where m.MOTCONSU_ID = @m_id
			                        )
		                        )
		                     begin
			                    set @ev_id=(SELECT TOP 1 MOTCONSU.MOTCONSU_EV_ID
			                    FROM PATDIREC PATDIREC 
			                    JOIN MOTCONSU MOTCONSU ON MOTCONSU.MOTCONSU_ID = PATDIREC.MOTCONSU_ID 
			                    JOIN PL_EXAM PL_EXAM ON PL_EXAM.PL_EXAM_ID = PATDIREC.PL_EXAM_ID
			                    WHERE PATDIREC.PATIENTS_ID= {1}   
				                    and datediff(day, PATDIREC.DATE_BIO,  @DATEtime)<= isnull(PL_EXAM.VAL_PERIOD, 2) 
                                order by PATDIREC.PATDIREC_ID)

                                exec up_get_id 'MOTCONSU',1,@m_id out
			                    INSERT into MOTCONSU (MOTCONSU_ID,    PATIENTS_ID,  MEDECINS_ID, FM_DEP_ID,MEDECINS_CREATE_ID, 
			                    MODELS_ID, DATE_CONSULTATION,  CREATE_DATE_TIME, REC_STATUS, MODIFY_DATE_TIME, MEDECINS_MODIFY_ID, KRN_MODIFY_USER_ID, 
			                    MOTCONSU_EV_ID, REC_NAME, MEDDEP_ID) 
			                    values (@m_id, {1} , {4}, {5}, {4}, {3}, {2}, @DATEtime,'W', @DATEtime, {4}, {4}, 
			                    @ev_id, {7}, {6})
		                    end

		                 --Обновление IMPDATA
		                 update IMPDATA set STATE=1
		                 where Impdata_ID={0} 
		                 --Обновление DS_RESTESTS
		                 update DS_RESTESTS set MOTCONSU_ID = @m_id
		                 where Impdata_ID={0}
                         select isnull(@m_id,0) 
                     ", ImpdataId, p_id, DateTime.Now, Processing.Ini.ModelId, Processing.MedicinsId, Processing.DepId, Processing.MedDepId, Processing.Mesure
                     ).Single();
                }
            }
            else
            {
                motconsuId = db.ExecuteQuery<int>(@"
                        update DS_RESTESTS set MOTCONSU_ID = 
                            (select top 1 dr.MOTCONSU_ID
                            from IMPDATA imp
                            join DS_RESTESTS dr on imp.Impdata_ID = dr.IMPDATA_ID
                            where imp.Impdata_ID = {0})
                        where MOTCONSU_ID is null
                            and Impdata_ID = {0}
                        
                        select top 1 dr.MOTCONSU_id 
                            from DS_RESTESTS dr
                            join MOTCONSU m on dr.MOTCONSU_ID = m.MOTCONSU_ID
                        where dr.IMPDATA_ID = {0} 
	                        and dr.MOTCONSU_ID is not null
                        ", ImpdataId).SingleOrDefault();

                if (p_id != 0 && motconsuId ==0)
                {
                    motconsuId = db.ExecuteQuery<int>(@"
                           	declare @m_id int, @ev_id int, @DATEtime datetime 
                            set @DATEtime = Getdate()
                            
		                    set	@m_id=(select top 1 m.MOTCONSU_ID from MOTCONSU m
		                    where PATIENTS_ID={1} 
		                    and convert (varchar(6), m.DATE_CONSULTATION, 12 ) = convert (varchar(6), {2}, 12 )
		                    and m.MODELS_ID={3} and m.MEDECINS_ID={4} and m.MEDDEP_ID = {6}  )
		
		                   if @m_id is null 
                              or exists 
		                        (select drImp.LAB_METHODS_ID 
		                        from IMPDATA imp
		                        join DS_RESTESTS drImp on imp.Impdata_ID = drImp.IMPDATA_ID
		                        where imp.Impdata_ID = {0} 
				                        and drImp.LAB_METHODS_ID = 
			                        ANY (
			                        select dr.LAB_METHODS_ID 
			                        from MOTCONSU m
			                        join DS_RESTESTS dr on m.MOTCONSU_ID = dr.MOTCONSU_ID
			                        where m.MOTCONSU_ID = @m_id
			                        )
		                        )
		                     begin
			                    set @ev_id=(SELECT TOP 1 MOTCONSU.MOTCONSU_EV_ID
			                    FROM PATDIREC PATDIREC 
			                    JOIN MOTCONSU MOTCONSU ON MOTCONSU.MOTCONSU_ID = PATDIREC.MOTCONSU_ID 
			                    JOIN PL_EXAM PL_EXAM ON PL_EXAM.PL_EXAM_ID = PATDIREC.PL_EXAM_ID
			                    WHERE PATDIREC.PATIENTS_ID= {1}   
				                    and datediff(day, PATDIREC.DATE_BIO,  @DATEtime)<= isnull(PL_EXAM.VAL_PERIOD, 2) 
                                order by PATDIREC.PATDIREC_ID)

                                exec up_get_id 'MOTCONSU',1,@m_id out
			                    INSERT into MOTCONSU (MOTCONSU_ID,    PATIENTS_ID,  MEDECINS_ID, FM_DEP_ID,MEDECINS_CREATE_ID, 
			                    MODELS_ID, DATE_CONSULTATION,  CREATE_DATE_TIME, REC_STATUS, MODIFY_DATE_TIME, MEDECINS_MODIFY_ID, KRN_MODIFY_USER_ID, 
			                    MOTCONSU_EV_ID, REC_NAME, MEDDEP_ID) 
			                    values (@m_id, {1} , {4}, {5}, {4}, {3}, {2}, @DATEtime,'W', @DATEtime, {4}, {4}, 
			                    @ev_id, {7}, {6})
		                    end

		                 --Обновление IMPDATA
		                 update IMPDATA set STATE=1
		                 where Impdata_ID={0} 
		                 --Обновление DS_RESTESTS
		                 update DS_RESTESTS set MOTCONSU_ID = @m_id
		                 where Impdata_ID={0}
                         select isnull(@m_id,0)
                     ", ImpdataId, p_id, resdate, Processing.Ini.ModelId, Processing.MedicinsId, Processing.DepId, Processing.MedDepId, Processing.Mesure
                     ).Single();
                }
            }

            db.Dispose();
            return motconsuId;
        }

        //Отправка в историю болезни (создание записей в таблице MOTCONSU) с обработкой ошибок
        static public void ErrInsertToMotconsu(int ImpdataId, int p_id, DateTime resdate, bool NewImport)
        {
            Console.WriteLine("Запись в истории болезни, ID пациента: {0}", p_id);
            try
            {
                InsertToMotconsu(ImpdataId, p_id, resdate, NewImport);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при вставке в MOTCONSU: {0}\n повторная попытка...", ex.Message);
                Files.WriteLogFile(string.Format("Ошибка при вставке в MOTCONSU: {0}", ex.ToString()));
                try { InsertToMotconsu(ImpdataId, p_id, resdate, NewImport); }
                catch (Exception exp)
                {
                    Console.WriteLine("Ошибка при вставке в MOTCONSU: {0}", exp.Message);
                    Files.WriteLogFile(string.Format("Ошибка при вставке в MOTCONSU: {0}", exp.ToString()));
                }
            }
        }

        static public int InsertToEMCe(int ImpdataId, int p_id, DateTime resdate, bool NewImport)
        {
            Console.WriteLine("Запись в истории болезни, ID пациента: {0}", p_id);
            int motconsuId = 0;
            try
            {
                motconsuId = InsertToEMC(ImpdataId, p_id, resdate, NewImport);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при вставке в MOTCONSU: {0}\n повторная попытка...", ex.Message);
                Files.WriteLogFile(string.Format("Ошибка при вставке в MOTCONSU: {0}", ex.ToString()));
                try { motconsuId = InsertToEMC(ImpdataId, p_id, resdate, NewImport); }
                catch (Exception exp)
                {
                    Console.WriteLine("Ошибка при вставке в MOTCONSU: {0}", exp.Message);
                    Files.WriteLogFile(string.Format("Ошибка при вставке в MOTCONSU: {0}", exp.ToString()));
                }
            }
            return motconsuId;
        }

        //Прикрепление pdf к записи ЭМК
        static public void AttachPdf(int motconsuId, int patientId, string filename, DateTime resDate, string folder, int rubricsId)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int state = db.ExecuteCommand(@"
                            declare @img_id int
                            exec up_get_id 'IMAGES', 1, @img_id out

                            insert IMAGES(Images_ID, PATIENTS_ID, Rubrics_ID, Date_Consultation, MOTCONSU_ID
                                , Descriptor, FileName, VIRTUAL_DISKS_ID, MEDECINS_CREATE_ID, FOLDER)
                            select @img_id, {0}, Rubrics_ID, {2}, {3}
		                            , {4}, {4}, VIRTUAL_DISKS_ID, {5}, {6}
                            from RUBRICS
                            where Rubrics_ID = {1}
                ", patientId, 16, resDate, motconsuId, filename, Processing.MedicinsId, folder);
        }

        static public void AttachPdfV2(int motconsuId, int patientId, string filename, DateTime resDate, string folder, int rubricsId)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int state = db.ExecuteCommand(@"
                            declare @img_id int
                            exec up_get_id 'IMAGES', 1, @img_id out

                            insert IMAGES(Images_ID, PATIENTS_ID, Rubrics_ID, Date_Consultation, MOTCONSU_ID
                                , Descriptor, FileName, VIRTUAL_DISKS_ID, MEDECINS_CREATE_ID, FOLDER)
                            select @img_id, {0}, Rubrics_ID, {2}, {3}
		                            , {4}, {4}, VIRTUAL_DISKS_ID, {5}, {6}
                            from RUBRICS
                            where Rubrics_ID = {1}
                ", patientId, rubricsId, resDate, motconsuId, filename, Processing.MedicinsId, folder);
        }

        static public void DeletePdfRefrense(int patientId, string filename)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int state = db.ExecuteCommand(@"delete from IMAGES
                                            where Patients_ID = {0}
                                                and Rubrics_ID = {1}
                                                and FileName = {2}
                 ", patientId, Processing.Ini.RubricsId, filename);
        }

        //Получение параметров записи по pdf файлу
        static public Processing.AttachData GetAttachData(string FileNamePart)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            Processing.AttachData aData = new Processing.AttachData();
            aData = db.ExecuteQuery<Processing.AttachData>(@"
                    select  top 1 i.PATIENTS_ID PatientId, m.MOTCONSU_ID MotconsuId, m.DATE_CONSULTATION ResDate
                    from IMPDATA i
                    join DS_RESTESTS r on i.Impdata_ID = r.IMPDATA_ID
                    join MOTCONSU m on r.MOTCONSU_ID = m.MOTCONSU_ID
                    where i.KEYCODE like '%' + {0} + '%'
                ", FileNamePart).SingleOrDefault();
            db.Dispose();
            return aData;
        }

        static public Processing.AttachData GetAttachDataS(string FileNamePart)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            Processing.AttachData aData = new Processing.AttachData();
            aData = db.ExecuteQuery<Processing.AttachData>(@"
                    select  top 1 i.PATIENTS_ID PatientId, m.MOTCONSU_ID MotconsuId, m.DATE_CONSULTATION ResDate
                    from IMPDATA i
                    join DS_RESTESTS r on i.Impdata_ID = r.IMPDATA_ID
                    join MOTCONSU m on r.MOTCONSU_ID = m.MOTCONSU_ID
                    where i.MSS_LAB_CODE like '%' + {0} + '%'
                        and DATEDIFF( d, m.DATE_CONSULTATION,GETDATE() ) < 30
                ", FileNamePart).SingleOrDefault();
            db.Dispose();
            return aData;
        }

        public class AnswerOrder
        {
            public int OrderId;
            public int PatdirecId;
            public int DirAnswId;
            public int MotconsuId;
            public int EventId;
        }
        
        //Отвт на заказ: Обновление DirAnsw 
        static public void UpdateDirAnsw(int ImpdataId)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            List<AnswerOrder> answOrder = db.ExecuteQuery<AnswerOrder>(@"
                    select ord.MSS_ORDER_REQUEST_ID OrderId, pd.PATDIREC_ID PatdirecId
	                    ,dan.DIR_ANSW_ID DirAnswId, dr.MOTCONSU_ID MotconsuId, isnull(m.MOTCONSU_EV_ID,0) EventId
                    from IMPDATA imp
                    join MSS_ORDER_REQUEST ord on imp.KEYCODE = ord.RequestCode
                    join DIR_SERV dsr on ord.MSS_ORDER_REQUEST_ID = dsr.MSS_ORDER_REQUEST_ID
                    join PATDIREC pd on dsr.PATDIREC_ID = pd.PATDIREC_ID
                    JOIN DIR_ANSW dan ON pd.PATDIREC_ID = dan.PATDIREC_ID 
                    LEFT OUTER JOIN MOTCONSU m ON m.MOTCONSU_ID = pd.MOTCONSU_ID 
                    JOIN DS_RESTESTS dr on imp.Impdata_ID = dr.IMPDATA_ID
                    where imp.Impdata_ID = {0}
                    group by ord.MSS_ORDER_REQUEST_ID, pd.PATDIREC_ID, dan.DIR_ANSW_ID, m.MOTCONSU_EV_ID, dr.MOTCONSU_ID
                ", ImpdataId).ToList();

            var gAnswOrder = from x in answOrder
                             group x by x.OrderId
                             into grouped
                             select new { grouped.Key, grouped };
            foreach(var g in gAnswOrder)
            {
                int state = db.ExecuteCommand(@"
                    update MSS_ORDER_REQUEST 
	                    set STATE = 3
                    where MSS_ORDER_REQUEST_ID = {0}                    
                    ", g.Key);

                foreach (var x in g.grouped)
                {
                    state = db.ExecuteCommand(@"
                        update DIR_ANSW 
                        set MOTCONSU_RESP_ID={1}
	                        , ANSW_STATE = 1
                        where DIR_ANSW_ID = {0}

                        if ( {1} <> 0)
                        update MOTCONSU 
                        set MOTCONSU_EV_ID = {2}
                        where MOTCONSU_ID = {1}
                        ",  x.DirAnswId, x.MotconsuId, x.EventId);
                }
            }
        }

        //Удаление ранее импортированных результатов если есть
        static public void RemoveDublicateResult(int ImpdataId)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int rowsAffected = db.ExecuteCommand(@"
                    declare @count int
	                set @count =(SELECT  COUNT(*) 
		                FROM IMPDATA impdata
		                join IMPDATA imp on imp.KEYCODE = impdata.KEYCODE 
				                and convert (varchar(6), imp.Date_Consultation, 12 ) = convert (varchar(6), IMPDATA.Date_Consultation, 12 )
				                and imp.Mesure = impdata.Mesure 
		                where impdata.Impdata_ID = {0} )

	                if @count > 1 
	                begin
		                update DS_RESTESTS set STATE = -1 
		                where Impdata_ID in ( select top (@count-1) imp.Impdata_ID
			                FROM IMPDATA impdata
			                join IMPDATA imp on imp.KEYCODE = impdata.KEYCODE 
					                and convert (varchar(6), imp.Date_Consultation, 12 ) = convert (varchar(6), IMPDATA.Date_Consultation, 12 )
					                and imp.Mesure = impdata.Mesure
		                where impdata.Impdata_ID = {0}
		                order by imp.Impdata_ID )
	                end
                    ", ImpdataId);
            db.Dispose();
        }

        //Отправка в историю болезни (создание записей в таблице MOTCONSU)
        static public void InsertToMotconsu(int ImpdataId, int p_id, DateTime resdate)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            if (p_id != 0)
            {
                int rowsAffected = db.ExecuteCommand(@"
                           	declare @m_id int, @ev_id int, @DATEtime datetime 
                            set @DATEtime = Getdate()
                            
		                    set	@m_id = null
                            /*
                              ( select top 1 m.MOTCONSU_ID 
                                from MOTCONSU m
                            		    join DS_RESTESTS dr on m.MOTCONSU_ID = dr.MOTCONSU_ID
	                                    left outer join DS_RESTESTS newdr 
	                                    left outer join IMPDATA imp on newdr.IMPDATA_ID = imp.Impdata_ID 
		                                    on dr.LAB_METHODS_ID = newdr.LAB_METHODS_ID and m.PATIENTS_ID = imp.PATIENTS_ID
			                                and newdr.IMPDATA_ID = {0} 
		                        where m.PATIENTS_ID={1} 
		                            and convert (varchar(6), m.DATE_CONSULTATION, 12 ) = convert (varchar(6), {2}, 12 )
		                            and m.MODELS_ID={3} and m.MEDECINS_ID={4}
                                    and newdr.DS_RESTESTS_ID is null
                              )
		                   */
		                   if @m_id is null 
		                     begin
			                    set @ev_id=(SELECT TOP 1 MOTCONSU.MOTCONSU_EV_ID
			                    FROM PATDIREC PATDIREC 
			                    JOIN MOTCONSU MOTCONSU ON MOTCONSU.MOTCONSU_ID = PATDIREC.MOTCONSU_ID 
			                    JOIN PL_EXAM PL_EXAM ON PL_EXAM.PL_EXAM_ID = PATDIREC.PL_EXAM_ID
			                    WHERE PATDIREC.PATIENTS_ID= {1}   
				                    and datediff(day, PATDIREC.DATE_BIO,  @DATEtime)<= PL_EXAM.VAL_PERIOD 
				                    and PL_EXAM.PL_EX_GR_ID=33
                                order by PATDIREC.PATDIREC_ID)

                                exec up_get_id 'MOTCONSU',1,@m_id out
			                    INSERT into MOTCONSU (MOTCONSU_ID,    PATIENTS_ID,  MEDECINS_ID, FM_DEP_ID,MEDECINS_CREATE_ID, 
			                    MODELS_ID, DATE_CONSULTATION,  CREATE_DATE_TIME, REC_STATUS, MODIFY_DATE_TIME, MEDECINS_MODIFY_ID, KRN_MODIFY_USER_ID, 
			                    MOTCONSU_EV_ID, REC_NAME, MEDDEP_ID) 
			                    values (@m_id, {1} , {4}, {5}, {4}, {3}, {2}, @DATEtime,'W', @DATEtime, {4}, {4}, 
			                    @ev_id, {7}, {6})
		                    end
		                 --Обновление IMPDATA
		                 update IMPDATA set STATE=1
		                 where Impdata_ID={0} 
		                 --Обновление DS_RESTESTS
		                 update DS_RESTESTS set MOTCONSU_ID = @m_id
		                 where Impdata_ID={0}"
                 , ImpdataId, p_id, resdate, Processing.Ini.ModelId, Processing.MedicinsId, Processing.DepId, Processing.MedDepId, Processing.Mesure);
            }
            db.Dispose();
        }
        
        //Отправка в историю болезни (создание записей в таблице MOTCONSU) с обработкой ошибок
        static public void ErrInsertToMotconsu(int ImpdataId, int p_id, DateTime resdate)
        {
            Console.WriteLine("Запись в истории болезни, ID пациента: {0}", p_id);
            try
            {
                InsertToMotconsu(ImpdataId, p_id, resdate);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при вставке в MOTCONSU: {0}\n повторная попытка...", ex.Message);
                Files.WriteLogFile(string.Format("Ошибка при вставке в MOTCONSU: {0}", ex.ToString()));
                try { InsertToMotconsu(ImpdataId, p_id, resdate); }
                catch (Exception exp)
                {
                    Console.WriteLine("Ошибка при вставке в MOTCONSU: {0}", exp.Message);
                    Files.WriteLogFile(string.Format("Ошибка при вставке в MOTCONSU: {0}", exp.ToString()));
                }
            }
        }

        //Вставка в таблицы ЭМК
        static public void InsetToUserTables(int ImpdataId)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int rowsAffected = db.ExecuteCommand(@"
                    declare @cnt int, @row_num int
                    set @row_num = 1
                    declare @user_table  table ( 
					                    row_num int
					                    ,motconsu_id int
					                    ,patient_id int
					                    ,table_name varchar(50)
					                    ,update_string nvarchar(max)
					                    )
                    insert @user_table
                    select ROW_NUMBER( ) OVER(order by p.TABLE_NAME)
		                    , m.motconsu_id,  m.PATIENTS_ID, p.TABLE_NAME
		                    ,  STUFF(cast ((select distinct [text()] = ',' + pp.FIELD_NAME + '=' + '''' + pr.VAL + ''''
		                    from DS_PARAMS pp 
		                    join  LAB_METHODBIO pb on pp.DS_PARAMS_ID=pb.DS_PARAMS_ID
		                    join DS_RESTESTS pr on pr.LAB_METHODS_ID = pb.LAB_METHODS_ID
		                    where pr.MOTCONSU_ID = r.MOTCONSU_ID
		                    for XML PATH(''), TYPE) as varchar(max)),1,1,'') 
                    from DS_RESTESTs r 
	                    join LAB_METHODBIO b on r.LAB_METHODS_ID = b.LAB_METHODS_ID
	                    join DS_PARAMS p on p.DS_PARAMS_ID=b.DS_PARAMS_ID
	                    join MOTCONSU m on m.MOTCONSU_ID=r.MOTCONSU_ID
                    where p.TABLE_NAME is not null 
	                    and FIELD_NAME is not null 
	                    and r.IMPDATA_ID = {0}
                    group by m.motconsu_id, m.PATIENTS_ID, p.TABLE_NAME, r.MOTCONSU_ID

                    declare @patient_id int, @motconsu_id int, @table_name varchar(50)
		                    , @update_string nvarchar(max), @insert_string nvarchar(max)
                    set @cnt = (select count(*) from @user_table)
                    while @row_num <= @cnt
                    begin
	                    select top 1  @patient_id = patient_id, @motconsu_id = motconsu_id
				                    , @table_name = table_name, @update_string = update_string
	                    from @user_table where row_num = @row_num

	                    set @insert_string = N'declare @dat_id int
		                    if not exists (select * from ' + @table_name + 
		                    ' where MOTCONSU_ID = @motconsu_id) 
		                    begin
		                     exec up_get_id ' + '''' + @table_name + '''' + ', 1, @dat_id out
		                     insert ' + @table_name + '(' + @table_name + '_ID, MOTCONSU_ID, PATIENTS_ID, DATE_CONSULTATION)
		                     values (@dat_id, @motconsu_id, @patient_id, GETDATE())
		                    end'
	                    exec sp_executesql @insert_string, 	N'@motconsu_id int, @patient_id int', @motconsu_id, @patient_id

	                    set @update_string=N'update ' + @table_name + ' SET ' + @update_string + 
	                     'where MOTCONSU_ID = @motconsu_id'
	                    exec sp_executesql @update_string, 	N'@motconsu_id int', @motconsu_id

	                    set @row_num = @row_num + 1
                    end
            ", ImpdataId);

            db.Dispose();
        }

        static public void InsertToEMCTables(int ImpdataId)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            var eMCTableData = db.ExecuteQuery<Processing.EMCTableData>(@"
                                                select distinct r.MOTCONSU_ID MotconsuId, m.PATIENTS_ID PatientId
                                                    , p.TABLE_NAME TableName, p.FIELD_NAME FieldName
                                                    , case 
                                                        when r.VAL = 'СМ.КОММ' or r.VAL = '.' then cast(r.COMMENT as varchar(max)) 
                                                        when r.VAL is not null then r.VAL 
                                                        when r.M_VAL is not null then cast(r.M_VAL as varchar(max)) 
                                                       end FieldValue
                                                from DS_RESTESTs r 
	                                                join LAB_METHODBIO b on r.LAB_METHODS_ID = b.LAB_METHODS_ID
	                                                join DS_PARAMS p on p.DS_PARAMS_ID=b.DS_PARAMS_ID
	                                                join MOTCONSU m on m.MOTCONSU_ID=r.MOTCONSU_ID
                                                where p.TABLE_NAME is not null 
	                                                and FIELD_NAME is not null 
	                                                and r.IMPDATA_ID = {0}
                                                ", ImpdataId);

            var gData = from x in eMCTableData 
                      group x by new
                      {
                          x.MotconsuId,
                          x.PatientId,
                          x.TableName
                      } into grouped
                      select new { grouped.Key, grouped };

            foreach (var g in gData)
            {
                Console.WriteLine("Вставка в таблицу ЭМК: {0}", g.Key.TableName);
                int state = db.ExecuteCommand(@"
                        declare @insert_string nvarchar(max)
                        set @insert_string = N'
                            declare @id int
		                    if not exists (select * from ' + {0} + 
		                    ' where MOTCONSU_ID = @motconsu_id) 
		                    begin
		                     exec up_get_id ' + '''' + {0} + '''' + ', 1, @id out
		                     insert ' + {0} + '(' + {0} + '_ID, MOTCONSU_ID, PATIENTS_ID, DATE_CONSULTATION)
		                     values (@id, @motconsu_id, @patient_id, GETDATE())
		                    end'
	                    exec sp_executesql @insert_string, 	N'@motconsu_id int, @patient_id int', {1}, {2}
                        ", g.Key.TableName, g.Key.MotconsuId, g.Key.PatientId);
       
                foreach (var item in g.grouped)
                {
                    state = db.ExecuteCommand(@"
 	                    declare @update_string nvarchar(max)
	                    set @update_string=N'update ' + {0} + ' SET ' + {2} + '=' + '''' + {3} + '''' +
	                     ' where MOTCONSU_ID = @motconsu_id'
	                    exec sp_executesql @update_string, 	N'@motconsu_id int', {1}
                        ", g.Key.TableName, g.Key.MotconsuId, item.FieldName, item.FieldValue);
                }
            }
 
            db.Dispose();
        }

        public class InOrder
        {
            public int OrdNum;
            public int PatientId;
            public int DirServId;
            public string BioTypeCode;
            public string ProfileCode;
            public int FilialId;
        }

        public class InfoOrder
        {
            public int PatientId;
            public int BioTypeId;
            public string BioTypeCode;
            public string ProfileCode;
            public int DirServId;
            public int FilialId;
        }

        //Получение номера закза
        public static int GetOrderNumber(DataContext db, int FilialId)
        {
            int ordNum = db.ExecuteQuery<int>(@"
                        declare @order_counter varchar(20), @ord_num int                 
                        set @order_counter='order_num'+'_'+  cast({0} as varchar(5))
                        declare  @t varchar(100)
                        exec up_get_counter_value @order_counter, 1, @ord_num out , @t out
                        select @ord_num
                   ", Processing.Ini.OrgId
                   ).SingleOrDefault();

            return ordNum;
        }
        
        public static string GetOrderCode(int OrderNum)
        {
            return DateTime.Now.ToString("yyMMdd", CultureInfo.InvariantCulture) + OrderNum.ToString("D4");
        }

        //Счетчик: получение значения
        public static int GetCounterValue(DataContext db, string counterName, int shift)
        {
            int lastvalue = db.ExecuteQuery<int>(@"
                                declare @val int, @t varchar(100)
                                exec up_get_counter_value {0}, {1}, @val out , @t out
                                select @val
                            ", counterName, shift).Single();
            return lastvalue;
        }

        public static int GetCounterValue(string counterName, int shift)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int lastvalue = db.ExecuteQuery<int>(@"
                                declare @val int, @t varchar(100)
                                exec up_get_counter_value {0}, {1}, @val out , @t out
                                select @val
                            ", counterName, shift).Single();
            return lastvalue;
        }

        //Формирование заказов 
        static public void CreateOrdersV2(out List<Request> RequestList)
        {
            //Создание контекста для работы с базой данных
            DataContext db = new DataContext(Processing.sqlConnectionStr);

            //Выбор информации для заказов во внешнюю лабораторию 
            Console.WriteLine("Информация для составления заказов");

            var InfoOrderList = db.ExecuteQuery<InfoOrder>(@"
                select distinct pd.PATIENTS_ID PatientId, lbt.LAB_BIOTYPE_ID BioTypeId
                        , lbt.CODE BioTypeCode, pr.CODE ProfileCode, dsr.DIR_SERV_ID  DirServId
                        , isnull(o.FM_ORG_MAIN_ID, o.FM_ORG_ID) FilialId
				from PATDIREC pd
                join FM_DEP d on pd.MEDECINS_BIO_DEP_ID = d.FM_DEP_ID
                join FM_ORG o on d.MAIN_ORG_ID = o.FM_ORG_ID
	            join DIR_SERV dsr on dsr.PATDIREC_ID=pd.PATDIREC_ID
	            left outer join MSS_ORDER_SERV osr on osr.FM_SERV_ID=dsr.FM_SERV_ID
	            join MSS_PROFILE_SERV prsr on prsr.FM_SERV_ID=dsr.FM_SERV_ID
	            join MSS_PROFILES pr on pr.MSS_PROFILES_ID = prsr.MSS_PROFILES_ID 
			            and (osr.FM_ORG_ID = pr.FM_ORG_ID or osr.FM_ORG_ID is null)
			    join MSS_PROFILES_BIOMATERIAL prbm on pr.MSS_PROFILES_ID = prbm.MSS_PROFILES_ID
											and ISNULL(prbm.NOT_USE,0) = 0
			    join MSS_BARCODE_NUMBERS brc on brc.PATDIREC_ID=pd.PATDIREC_ID
	            join LAB_BIOTYPE lbt on brc.LAB_BIOTYPE_ID = lbt.LAB_BIOTYPE_ID
						and prbm.LAB_BIOTYPE_ID = lbt.LAB_BIOTYPE_ID
	            left outer join MSS_ORDER_REQUEST ord on dsr.MSS_ORDER_REQUEST_ID = ord.MSS_ORDER_REQUEST_ID
                left outer join FM_ORG  eo on dsr.FM_ORG_ID = eo.FM_ORG_ID
	            where  datediff(HH,pd.DATE_BIO, GETDATE()) <= 24 
	                and datediff(mi,pd.DATE_BIO, GETDATE()) > 5
	                and ord.MSS_ORDER_REQUEST_ID is null 
                    and ISNULL(osr.state,0) = 1
                    and  ISNULL(eo.FM_ORG_ID,0) = 0 
                    and pr.FM_ORG_ID = {0}  
                ", Processing.Ini.OrgId);

            //Группировака по пациентам
            var gInfoOrder = from x in InfoOrderList
                             group x by x.PatientId;

            List<InOrder> inOrderList = new List<InOrder>();

            foreach (var g in gInfoOrder)
            {
                 //Распределение профилей по заказам 
                foreach (var x in g)
                {
                    if (Processing.Ini.DriverCode == "KDL")
                    {
                        //Подсчет количества разных биоматериалов для каждого повторяющегося профиля
                        int OrderNum = (from xq in inOrderList
                                        where xq.ProfileCode == x.ProfileCode
                                                && xq.BioTypeCode != x.BioTypeCode
                                        select new
                                        {
                                            xq.ProfileCode,
                                            xq.BioTypeCode
                                        }).Distinct().Count();

                        for (int i = 0; i < 20; i++)
                        {
                            if (OrderNum == i)
                                inOrderList.Add(
                                    new InOrder
                                    {
                                        OrdNum = i + 1,
                                        PatientId = g.Key,
                                        DirServId = x.DirServId,
                                        ProfileCode = x.ProfileCode,
                                        BioTypeCode = x.BioTypeCode
                                    });
                        }
                    }
                    else
                    {
                        inOrderList.Add(
                                    new InOrder
                                    {
                                        OrdNum = 1,
                                        PatientId = g.Key,
                                        DirServId = x.DirServId,
                                        ProfileCode = x.ProfileCode,
                                        BioTypeCode = x.BioTypeCode,
                                        FilialId = x.FilialId
                                    });
                    }
                }
            }

            //Группировка по заказам
            var Orders = from x in inOrderList
                        group x by new
                        {
                            x.PatientId,
                            x.OrdNum,
                            x.FilialId
                        };
            //Создание заказов           
            foreach (var order in Orders)
            {
                Console.WriteLine("Создание заказа {1} для пациента {0}", order.Key.PatientId, order.Key.OrdNum);

                int orderNumber = GetOrderNumber(db, 0);
                string orderCode = DateTime.Now.ToString("yyMMdd") + orderNumber.ToString("D3");
                if (Processing.Ini.DriverCode == "KDL")   orderCode = orderNumber.ToString();
                //Создание строки заказа в MSS_REQUEST_ORDER
                int orderId = db.ExecuteQuery<int>(@"
                        declare @ord_id int, @date datetime
                        set @date = GETDATE()
                        exec up_get_id 'MSS_ORDER_REQUEST',1, @ord_id out	
                        
                        insert MSS_ORDER_REQUEST (MSS_ORDER_REQUEST_ID, ORDER_NUMBER, FM_ORG_ID, PATIENTS_ID, DATE_ORDER, STATE, RequestCode)
                        values (@ord_id, {2}, {0}, {1}, @date, 0, {3})
                        select @ord_id
                    ", Processing.Ini.OrgId, order.Key.PatientId, orderNumber, orderCode).SingleOrDefault();

                //Создание состава заказа: Обновление DIR_SERV 
                string idlist = string.Empty;
                foreach (var o in order)
                {
                    idlist += ", " + o.DirServId.ToString();
                }

                int ok = db.ExecuteCommand(@"
		                declare @update_string nvarchar(max)
	                    set @update_string = N'update DIR_SERV set MSS_ORDER_REQUEST_ID = ' + cast({1} as nvarchar(6)) + '
	                    where DIR_SERV_ID in (' + {0} + ')'
	                    exec sp_executesql @update_string
                            ", idlist.Substring(2), orderId);

                //Обновление штрих-кодов биоматериала
                if (Processing.Ini.DriverCode == "KDL")
                {
                    if (Processing.Ini.BarcodeOption != "0")
                    {
                        var BmTypes = from x in order
                                      group x by x.BioTypeCode;
                        int i = 1;
                        foreach (var bmtype in BmTypes)
                        {
                            string barcode;
                            switch (Processing.Ini.BarcodeOption)
                            {
                                case "1": barcode = orderCode; break;
                                case "2": barcode = orderCode + i.ToString(); break;
                                case "3": barcode = orderCode + i.ToString("D2"); break;
                                default: barcode = string.Empty; break;
                            }

                            ok = db.ExecuteCommand(@"
		                        declare @update_string nvarchar(max)
	                            set @update_string = N'update MSS_BARCODE_NUMBERS set BARCODE_NUMBER = ''' + {2} + ''' 
							        from MSS_BARCODE_NUMBERS brc
							        join DIR_SERV dsr on brc.PATDIREC_ID = dsr.PATDIREC_ID
							        join LAB_BIOTYPE lbt on brc.LAB_BIOTYPE_ID = lbt.LAB_BIOTYPE_ID
							        where lbt.CODE = ''' + {1} + '''
								        and dsr.DIR_SERV_ID in ( ' + {0} + ' )'
	                            exec sp_executesql @update_string
                            ", idlist.Substring(2), bmtype.Key, barcode);
                            i++;
                        }
                    }
                }

                
                if (Processing.Ini.DriverCode == "SLS")
                {
                     var BmTypes = from x in order
                                  group x by x.BioTypeCode;

                    string counterName = string.Format("LAB_{0}", Processing.Ini.OrgId);      //string.Format("LAB_{0}_{1}", Processing.Ini.OrgId, order.Key.FilialId);

                    foreach (var bmtype in BmTypes)
                    {
                        string barcode = DateTime.Now.ToString("yyMMdd") + GetCounterValue(db, counterName, 1).ToString("D4");

                        ok = db.ExecuteCommand(@"
		                        declare @update_string nvarchar(max)
	                            set @update_string = N'update MSS_BARCODE_NUMBERS set BARCODE_NUMBER = ''' + {2} + ''' 
							        from MSS_BARCODE_NUMBERS brc
							        join DIR_SERV dsr on brc.PATDIREC_ID = dsr.PATDIREC_ID
							        join LAB_BIOTYPE lbt on brc.LAB_BIOTYPE_ID = lbt.LAB_BIOTYPE_ID
							        where lbt.CODE = ''' + {1} + '''
								        and dsr.DIR_SERV_ID in ( ' + {0} + ' )'
	                            exec sp_executesql @update_string
                            ", idlist.Substring(2), bmtype.Key, barcode);
                    }
                }
            }

            //Заказы
            RequestList = db.ExecuteQuery<Request>(@"
                          select distinct req.MSS_ORDER_REQUEST_ID ID, req.RequestCode Code, req.ORDER_NUMBER Nr
								, req.DATE_ORDER Date, req.PATIENTS_ID PatientID
                                , isnull(oe.Code,'') FilialCode, o.LABEL FilialLabel, isnull(oe.WEB_PSWD,'') FilialPass
                                , isnull(oe.DEP_Code,'') DepCode, isnull(oe.DEP_Label,'') DepLabel
                            from PATDIREC pd
                            join DIR_SERV dsr on pd.PATDIREC_ID = dsr.PATDIREC_ID
                            join MSS_ORDER_REQUEST req on dsr.MSS_ORDER_REQUEST_ID = req.MSS_ORDER_REQUEST_ID
                            join FM_DEP d on pd.MEDECINS_BIO_DEP_ID = d.FM_DEP_ID
                            join FM_ORG o on d.MAIN_ORG_ID = o.FM_ORG_ID
                            left outer join mss_external_org_codes oe on oe.FM_ORG_ID = o.FM_ORG_ID 
										and oe.EXTERNAL_ORG_ID = {0}
							where req.STATE = 0 
								and datediff(HH, req.DATE_ORDER, GETDATE()) <= 48
                            ", Processing.Ini.OrgId).ToList();

            if (Processing.Ini.DriverCode == "SLS")
            {
                foreach (var qrequest in RequestList)
                {
                    qrequest.FilialPass = Crypto.DecryptStringAES(Crypto.DStr(qrequest.FilialPass), "hip-hop-lab");
                }
            }

            db.Dispose();
        }

        //Получение списка заказов
        static public void GetOrdersList(out List<Request> RequestList)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);

            //Выбор заказов готовых к отправке 
            Console.WriteLine("Получение списка готовых заказов");
            RequestList = db.ExecuteQuery<Request>(@"
                          select distinct req.MSS_ORDER_REQUEST_ID ID, req.RequestCode Code, req.ORDER_NUMBER Nr
								, req.DATE_ORDER Date, req.PATIENTS_ID PatientID
                                , isnull(oe.Code,'') FilialCode, o.LABEL FilialLabel, isnull(oe.WEB_PSWD,'') FilialPass
                                , isnull(oe.DEP_Code,'') DepCode, isnull(oe.DEP_Label,'') DepLabel
                                , oe.WEB_ID WebId
                            from PATDIREC pd
                            join DIR_SERV dsr on pd.PATDIREC_ID = dsr.PATDIREC_ID
                            join MSS_ORDER_REQUEST req on dsr.MSS_ORDER_REQUEST_ID = req.MSS_ORDER_REQUEST_ID
                            join FM_DEP d on pd.MEDECINS_BIO_DEP_ID = d.FM_DEP_ID
                            join FM_ORG o on d.MAIN_ORG_ID = o.FM_ORG_ID
                            left outer join mss_external_org_codes oe on oe.FM_ORG_ID = o.FM_ORG_ID 
										and oe.EXTERNAL_ORG_ID = {0}
							where req.STATE = 0 
								and datediff(HH, req.DATE_ORDER, GETDATE()) <= 24
                                and datediff(mi,pd.DATE_BIO, GETDATE())> = 0
                            ", Processing.Ini.OrgId).ToList();

            if (Processing.Ini.DriverCode == "SLS")
            {
                foreach (var qrequest in RequestList)
                {
                    qrequest.FilialPass = Crypto.DecryptStringAES(Crypto.DStr(qrequest.FilialPass), "hip-hop-lab");
                }
            }

            db.Dispose();
        }
        static public int execTest()
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int cnt = db.ExecuteQuery<int>(@"select count(*) from mss_order_request where DATE_ORDER > '2022-07-21T00:00:00.000'").First();
            db.Dispose();
            return cnt;
        }

        //Удаление заказа
        static public void DeleteOrder(string ExtOrderId)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);

            //Выбор заказов готовых к отправке 
            Console.WriteLine("Удаление заказов");
            int status = db.ExecuteCommand (@"
                         delete brc
						 from MSS_BARCODE_NUMBERS brc
						 join MSS_ORDER_REQUEST ord on ord.MSS_ORDER_REQUEST_ID = brc.MSS_ORDER_REQUEST_ID
					     where ord.EXTERNAL_ID = {0}

                         delete from MSS_ORDER_REQUEST
                         where TO_DELETE = 1 
	                        and EXTERNAL_ID = {0}
                            ", ExtOrderId);
            db.Dispose();
        }

        //Список заказов на удаление
        static public List<string> GetOrdersToDelete()
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            List<string> deleteList = db.ExecuteQuery<string>(@"
                        select EXTERNAL_ID 
                        from MSS_ORDER_REQUEST
                        where state = -2"
                ).DefaultIfEmpty().ToList();
            db.Dispose();
            return deleteList;
        }

        //Получение списка заказов
        static public void GetOrdersList(int State, int bdelay, out List<Request> RequestList)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);

            //Выбор заказов готовых к отправке 
            Console.WriteLine("Получение списка готовых заказов");
            RequestList = db.ExecuteQuery<Request>(@"
                          select req.MSS_ORDER_REQUEST_ID ID, req.RequestCode Code, req.ORDER_NUMBER Nr
								, req.DATE_ORDER Date, req.PATIENTS_ID PatientID
                                , isnull(oe.Code,'') FilialCode, o.LABEL FilialLabel, isnull(oe.WEB_PSWD,'') FilialPass
                                , case when fd2.FM_DEP_ID is null then isnull(d.CODE,'') else fd2.CODE end DepCode
								, case when fd2.FM_DEP_ID is null then isnull(d.Label,'') else fd2.Label end DepLabel
								, case when fd2.FM_DEP_ID is null then isnull(d.SHORT_NAME,'1') else fd2.SHORT_NAME end DepId
								, oe.WEB_ID WebId, MAX(cast(pd.CITO as int))
                                , stuff((select  distinct '; ' + cast(pdd.COMMENTAIRE as varchar(max))
									    from PATDIREC pdd
									    join DIR_SERV dsrd on pdd.PATDIREC_ID = dsrd.PATDIREC_ID
									    join MSS_ORDER_REQUEST reqd on dsrd.MSS_ORDER_REQUEST_ID = reqd.MSS_ORDER_REQUEST_ID
									    where reqd.MSS_ORDER_REQUEST_ID = req.MSS_ORDER_REQUEST_ID
									    for XML PATH(''), TYPE).value('.', 'varchar(max)'),1,2,'') Comment
                            from PATDIREC pd
							LEFT OUTER JOIN MEDDEP md2 ON md2.MEDDEP_ID = pd.MEDDEP_ID
							LEFT OUTER JOIN FM_DEP fd2 ON fd2.FM_DEP_ID = md2.FM_DEP_ID
                            join DIR_SERV dsr on pd.PATDIREC_ID = dsr.PATDIREC_ID
                            join MSS_ORDER_REQUEST req on dsr.MSS_ORDER_REQUEST_ID = req.MSS_ORDER_REQUEST_ID
                            join FM_DEP d on pd.MEDECINS_BIO_DEP_ID = d.FM_DEP_ID
                            join FM_ORG o on d.MAIN_ORG_ID = o.FM_ORG_ID
                            left outer join mss_external_org_codes oe on oe.FM_ORG_ID = o.FM_ORG_ID 
										and oe.EXTERNAL_ORG_ID = {0}
							where req.STATE = {1} 
								and datediff(HH, req.DATE_ORDER, GETDATE()) <= 120
                                and datediff(mi,pd.DATE_BIO, GETDATE()) >= {2}
                                and req.FM_ORG_ID = {0}
                            group by req.MSS_ORDER_REQUEST_ID, req.RequestCode, req.ORDER_NUMBER
								, req.DATE_ORDER, req.PATIENTS_ID
								, oe.Code, o.LABEL, oe.WEB_PSWD
                                , d.Code, d.Label, oe.WEB_ID, fd2.FM_DEP_ID, fd2.CODE, fd2.Label
                                , fd2.SHORT_NAME, d.SHORT_NAME
                            ", Processing.Ini.OrgId, State, bdelay).ToList();

            if (Processing.Ini.DriverCode == "SLS")
            {
                foreach (var qrequest in RequestList)
                {
                    //qrequest.FilialPass = Crypto.DecryptStringAES(Crypto.DStr(qrequest.FilialPass), "hip-hop-lab");
                }
            }

            db.Dispose();
        }

        static public void GetOrdersListV2(int State, int bdelay, out List<Request> RequestList)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);

            //Выбор заказов готовых к отправке 
            Console.WriteLine("Получение списка готовых заказов");
            RequestList = db.ExecuteQuery<Request>(@"
                          select req.MSS_ORDER_REQUEST_ID ID, req.RequestCode Code, req.ORDER_NUMBER Nr
								, req.DATE_ORDER Date, req.PATIENTS_ID PatientID
                                , isnull(oe.Code,'') FilialCode, o.LABEL FilialLabel, isnull(oe.WEB_PSWD,'') FilialPass
								, oe.WEB_ID WebId, MAX(cast(pd.CITO as int))
                                , stuff((select  distinct '; ' + cast(pdd.COMMENTAIRE as varchar(max))
									    from PATDIREC pdd
									    join DIR_SERV dsrd on pdd.PATDIREC_ID = dsrd.PATDIREC_ID
									    join MSS_ORDER_REQUEST reqd on dsrd.MSS_ORDER_REQUEST_ID = reqd.MSS_ORDER_REQUEST_ID
									    where reqd.MSS_ORDER_REQUEST_ID = req.MSS_ORDER_REQUEST_ID
									    for XML PATH(''), TYPE).value('.', 'varchar(max)'),1,2,'') Comment
                            from PATDIREC pd
							LEFT OUTER JOIN MEDDEP md2 ON md2.MEDDEP_ID = pd.MEDDEP_ID
							LEFT OUTER JOIN FM_DEP fd2 ON fd2.FM_DEP_ID = md2.FM_DEP_ID
                            join DIR_SERV dsr on pd.PATDIREC_ID = dsr.PATDIREC_ID
                            join MSS_ORDER_REQUEST req on dsr.MSS_ORDER_REQUEST_ID = req.MSS_ORDER_REQUEST_ID
                            join FM_DEP d on pd.MEDECINS_BIO_DEP_ID = d.FM_DEP_ID
                            join FM_ORG o on d.MAIN_ORG_ID = o.FM_ORG_ID
                            left outer join mss_external_org_codes oe on oe.FM_ORG_ID = o.FM_ORG_ID 
										and oe.EXTERNAL_ORG_ID = {0}
							where req.STATE = {1} 
								and datediff(HH, req.DATE_ORDER, GETDATE()) <= 120
                                and datediff(mi,pd.DATE_BIO, GETDATE()) >= {2}
                                and req.FM_ORG_ID = {0}
                            group by req.MSS_ORDER_REQUEST_ID, req.RequestCode, req.ORDER_NUMBER
								, req.DATE_ORDER, req.PATIENTS_ID
								, oe.Code, o.LABEL, oe.WEB_PSWD
                                , oe.WEB_ID
                            ", Processing.Ini.OrgId, State, bdelay).ToList();

            if (Processing.Ini.DriverCode == "SLS")
            {
                foreach (var qrequest in RequestList)
                {
                    //qrequest.FilialPass = Crypto.DecryptStringAES(Crypto.DStr(qrequest.FilialPass), "hip-hop-lab");
                }
            }

            db.Dispose();
        }
        //Формирование заказов 
        static public void CreateOrders(out List<Request> RequestList)
        {
            //Создание контекста для работы с базой данных
            DataContext db = new DataContext(Processing.sqlConnectionStr);

            //Выбор списка пациентов с направлениями во внешнюю лабораторию 
            Console.WriteLine("Создание списка пациентов для составления заказов");
            var PatientsIDList = db.ExecuteQuery<int>(@"
                select distinct pd.PATIENTS_ID PATIENTS_ID
                from PATDIREC pd
	            join DIR_SERV dsr on dsr.PATDIREC_ID=pd.PATDIREC_ID
	            left outer join MSS_ORDER_SERV osr on osr.FM_SERV_ID=dsr.FM_SERV_ID
	            join MSS_PROFILE_SERV prsr on prsr.FM_SERV_ID=dsr.FM_SERV_ID
	            join MSS_PROFILES pr on pr.MSS_PROFILES_ID=prsr.MSS_PROFILES_ID 
			            and (osr.FM_ORG_ID=pr.FM_ORG_ID or osr.FM_ORG_ID is null)
	            join MSS_BARCODE_NUMBERS brc on brc.PATDIREC_ID=pd.PATDIREC_ID
                left outer join MSS_ORDER_PATDIREC opd 
                left outer join MSS_ORDER_REQUEST ord on opd.MSS_ORDER_REQUEST_ID = ord.MSS_ORDER_REQUEST_ID 
                            on pd.PATDIREC_ID = opd.PATDIREC_ID and ord.FM_ORG_ID = pr.FM_ORG_ID
                left outer join FM_ORG  eo on dsr.FM_ORG_ID = eo.FM_ORG_ID
	            where  datediff(HH,pd.DATE_BIO, GETDATE())<=24 
	                and datediff(mi,pd.DATE_BIO, GETDATE())>5
	                and ord.MSS_ORDER_REQUEST_ID is null 
                    and isnull(osr.state,0)=1
                    and pr.FM_ORG_ID= {0}
                    and  (ISNULL(eo.FM_ORG_ID,0) = 0 or  eo.ORG_TYPE = 'I' )    
                ", Processing.Ini.OrgId);

            //Создание заказов в таблице MSS_ORDER_REQUEST
            RequestList = new List<Request>();
            foreach (int ID in PatientsIDList)
            {
                Console.WriteLine("Создание заказа для пациента ID = {0}", ID);
                Request qrequest = db.ExecuteQuery<Request>(
                    @" declare @order_counter varchar(20), @ord_num int, @ord_id int
                    declare @date datetime 
                    set @date = GETDATE()
                    exec up_get_id 'MSS_ORDER_REQUEST',1, @ord_id out	

	                set @order_counter='order_num'+'_'+  cast({1} as varchar(5))
                    declare  @t varchar(100)
                    exec up_get_counter_value @order_counter, 1, @ord_num out , @t out

                    declare @ord_str varchar(17) 
                    set @ord_str = cast(@ord_num as varchar(17))
                    declare @ext_code varchar(20) 
                    set @ext_code = substring(CONVERT(varchar, @date, 112),3,6) 
                            + substring(REPLICATE('0',4),0, 4-LEN(@ord_str)+1)+ @ord_str

	                insert MSS_ORDER_REQUEST (MSS_ORDER_REQUEST_ID, ORDER_NUMBER, FM_ORG_ID, PATIENTS_ID, DATE_ORDER, STATE, RequestCode)
	                values (@ord_id, @ord_num, {1}, {0}, @date, 0, @ext_code)

                    declare @pdt table (id int)
					insert @pdt
                    select distinct pd.PATDIREC_ID
                        from PATDIREC pd
	                    join DIR_SERV dsr on dsr.PATDIREC_ID=pd.PATDIREC_ID
	                    join MSS_ORDER_SERV osr on osr.FM_SERV_ID=dsr.FM_SERV_ID
	                    join MSS_BARCODE_NUMBERS brc on brc.PATDIREC_ID=pd.PATDIREC_ID and brc.FM_ORG_ID = {1}
	                    left outer join MSS_ORDER_PATDIREC opd on pd.PATDIREC_ID = opd.PATDIREC_ID
	                    left outer join MSS_ORDER_REQUEST ord on opd.MSS_ORDER_REQUEST_ID = ord.MSS_ORDER_REQUEST_ID 
                        left outer join FM_ORG  eo on dsr.FM_ORG_ID = eo.FM_ORG_ID
	                    where pd.PATIENTS_ID = {0} 
                            and datediff(HH,pd.DATE_BIO, GETDATE())<=24
                            and datediff(mi,pd.DATE_BIO, GETDATE())>5
						    and ord.MSS_ORDER_REQUEST_ID is null
		                    and isnull(osr.state,0)=1 
                            and (ISNULL(eo.FM_ORG_ID,0) = 0 or  eo.ORG_TYPE = 'I' ) 
                    
                    declare @pdcount int 
                    set @pdcount = (select COUNT(*) from @pdt) 
                    declare @ordpd_id int
                    exec up_get_id 'MSS_ORDER_PATDIREC', @pdcount, @ordpd_id out 

                    insert MSS_ORDER_PATDIREC (MSS_ORDER_PATDIREC_ID, PATDIREC_ID, MSS_ORDER_REQUEST_ID)
                    select ROW_NUMBER()OVER(order by id) - 1 +  @ordpd_id, id, @ord_id from @pdt

                    select distinct @ord_id ID, @ext_code Code,  @ord_num Nr, @date Date, {0} PatientID
                        , isnull(oe.Code,'') FilialCode, o.LABEL FilialLabel, isnull(oe.WEB_PSWD,'') FilialPass
                        , isnull(oe.DEP_Code,'') DepCode, isnull(oe.DEP_Label,'') DepLabel
                    from PATDIREC pd
                    join MSS_ORDER_PATDIREC opd on pd.PATDIREC_ID = opd.PATDIREC_ID
                    join FM_DEP d on pd.MEDECINS_BIO_DEP_ID = d.FM_DEP_ID
                    join FM_ORG o on d.MAIN_ORG_ID = o.FM_ORG_ID
                    left outer join mss_external_org_codes oe on oe.FM_ORG_ID = o.FM_ORG_ID and oe.EXTERNAL_ORG_ID = {1}
                    where opd.MSS_ORDER_REQUEST_ID = @ord_id
                ", ID, Processing.Ini.OrgId).Single();

                if (Processing.Ini.DriverCode == "SLS")
                {
                    UpdateBarCode(db, qrequest.ID);
                    qrequest.FilialPass = Crypto.DecryptStringAES(Crypto.DStr(qrequest.FilialPass), "hip-hop-lab");
                }

                RequestList.Add(qrequest);
            }

            foreach (Request req in RequestList) Console.WriteLine("Создан заказ № {0}", req.Code);

            db.Dispose();
        }

        //Проверка наличия забора биоматериала
        static public bool CheckOrders(out List<CheckInfo> CheckInfoList)
        {
            //Создание контекста для работы с базой данных
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            CheckInfoList = new List<CheckInfo>();
            //Выбор списка пациентов с направлениями во внешнюю лабораторию 
            CheckInfoList = db.ExecuteQuery<CheckInfo>(@"
                select distinct pd.PATIENTS_ID PatientId, pd.FM_INTORG_ID FilialId
                from PATDIREC pd
	            join DIR_SERV dsr on dsr.PATDIREC_ID=pd.PATDIREC_ID
	            left outer join MSS_ORDER_SERV osr on osr.FM_SERV_ID=dsr.FM_SERV_ID
	            join MSS_PROFILE_SERV prsr on prsr.FM_SERV_ID=dsr.FM_SERV_ID
	            join MSS_PROFILES pr on pr.MSS_PROFILES_ID=prsr.MSS_PROFILES_ID 
			            and (osr.FM_ORG_ID=pr.FM_ORG_ID or osr.FM_ORG_ID is null)
	            join MSS_BARCODE_NUMBERS brc on brc.PATDIREC_ID=pd.PATDIREC_ID
                left outer join MSS_ORDER_PATDIREC opd 
                left outer join MSS_ORDER_REQUEST ord on opd.MSS_ORDER_REQUEST_ID = ord.MSS_ORDER_REQUEST_ID 
                            on pd.PATDIREC_ID = opd.PATDIREC_ID and ord.FM_ORG_ID = pr.FM_ORG_ID
                left outer join FM_ORG  eo on dsr.FM_ORG_ID = eo.FM_ORG_ID
	            where  datediff(HH,pd.DATE_BIO, GETDATE())<=24 
	                and ord.MSS_ORDER_REQUEST_ID is null 
                    and isnull(osr.state,0)=1
                    and pr.FM_ORG_ID= {0}
                 ", Processing.Ini.OrgId).ToList();
            bool ch = (CheckInfoList.Count() != 0);
            if (ch) Console.WriteLine("Новые заказы!");
            db.Dispose();
            return ch;
        }

        static public bool CheckOrdersV2(out List<CheckInfo> CheckInfoList)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            CheckInfoList = new List<CheckInfo>();
            CheckInfoList = db.ExecuteQuery<CheckInfo>(@"
                    select distinct ord.MSS_ORDER_REQUEST_ID OrderId, ord.PATIENTS_ID PatientId, pd.FM_INTORG_ID FilialId
                    from MSS_ORDER_REQUEST ord
                    join DIR_SERV dsr on ord.MSS_ORDER_REQUEST_ID = dsr.MSS_ORDER_REQUEST_ID
                    join PATDIREC pd on dsr.PATDIREC_ID = pd.PATDIREC_ID
                    where isnull(ord.STATE,0) = 0
	                    and ord.FM_ORG_ID = 975
                    ", Processing.Ini.OrgId).ToList();
            bool ch = (CheckInfoList.Count() != 0);
            if (ch) Console.WriteLine("Новые заказы!");
            db.Dispose();
            return ch;
        }

        //Формирование заказов в реальном времени
        static public void CreateOrder(CheckInfo checkinfo, out List<Processing.RequestOrder> request)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);

            //Создание заказа в таблице MSS_ORDER_REQUEST
            Console.WriteLine("Создание заказа для пациента ID = {0}", checkinfo.PatientId);
            request = new List<Processing.RequestOrder>();
            request = db.ExecuteQuery<Processing.RequestOrder>(@"
                    declare @order_counter varchar(20), @ord_num int, @ord_id int
                    declare @date datetime = GETDATE()
                    exec up_get_id 'MSS_ORDER_REQUEST',1, @ord_id out	

	                set @order_counter='order_num'+'_' +  cast({1} as varchar(5))
	                declare  @t varchar(100)
                    exec up_get_counter_value @order_counter, 1, @ord_num out , @t out
                    declare @ord_str varchar(17) 
                    set @ord_str = cast(@ord_num as varchar(17))
                    declare @ext_code varchar(20) 
                    set @ext_code = substring(CONVERT(varchar, @date, 112),3,6) 
                            + substring(REPLICATE('0',4),0, 4-LEN(@ord_str)+1)+ @ord_str

	                insert MSS_ORDER_REQUEST (MSS_ORDER_REQUEST_ID, ORDER_NUMBER, FM_ORG_ID, PATIENTS_ID, DATE_ORDER, STATE, RequestCode)
	                values (@ord_id, @ord_num, {1}, {0}, @date, 0, @ext_code)

                    declare @pdt table (id int)
					insert @pdt
                    select distinct pd.PATDIREC_ID
                        from PATDIREC pd
	                    join DIR_SERV dsr on dsr.PATDIREC_ID=pd.PATDIREC_ID
	                    join MSS_ORDER_SERV osr on osr.FM_SERV_ID=dsr.FM_SERV_ID
	                    join MSS_BARCODE_NUMBERS brc on brc.PATDIREC_ID=pd.PATDIREC_ID and brc.FM_ORG_ID = {1}
	                    left outer join MSS_ORDER_PATDIREC opd on pd.PATDIREC_ID = opd.PATDIREC_ID
	                    left outer join MSS_ORDER_REQUEST ord on opd.MSS_ORDER_REQUEST_ID = ord.MSS_ORDER_REQUEST_ID 
                        left outer join FM_ORG  eo on dsr.FM_ORG_ID = eo.FM_ORG_ID
	                    where pd.PATIENTS_ID = {0} 
                            and datediff(HH,pd.DATE_BIO, GETDATE())<=24
						    and ord.MSS_ORDER_REQUEST_ID is null
		                    and isnull(osr.state,0)=1 

                    declare @pdcount int = (select COUNT(*) from @pdt) 
                    declare @ordpd_id int
                    exec up_get_id 'MSS_ORDER_PATDIREC', @pdcount, @ordpd_id out 

                    insert MSS_ORDER_PATDIREC (MSS_ORDER_PATDIREC_ID, PATDIREC_ID, MSS_ORDER_REQUEST_ID)
                    select ROW_NUMBER()OVER(order by id) - 1 +  @ordpd_id, id, @ord_id from @pdt

                    select distinct @ord_id RequestId, @ext_code RequestCode,  @ord_num RequestNr, @date RequestDate
                           , pd.FM_INTORG_ID FilialId, isnull(oe.Code,'') FilialCode, o.LABEL FilialLabel
                           , {0} PatientID, p.NOM LastName, p.PRENOM FirstName, p.PATRONYME MiddleName, p.POL Gender, p.NE_LE BirthDate
                           ,  pr.CODE ProductCode, pr.LABEL ProductName, brn.MSS_BARCODE_NUMBERS_ID BarcodeId
                           , prbm.OPTION_SET BiomaterialOption, bm.CODE BiomaterialCode
                    from PATDIREC pd
                    join MSS_ORDER_PATDIREC opd on pd.PATDIREC_ID = opd.PATDIREC_ID
					join MSS_ORDER_REQUEST ord on opd.MSS_ORDER_REQUEST_ID = ord.MSS_ORDER_REQUEST_ID
                    join MSS_BARCODE_NUMBERS brn on pd.PATDIREC_ID = brn.PATDIREC_ID and brn.FM_ORG_ID = ord.FM_ORG_ID
                    join LAB_BIOTYPE bm on brn.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
                    join DIR_SERV ds on pd.PATDIREC_ID = ds.PATDIREC_ID
                    join MSS_PROFILE_SERV ps on ds.FM_SERV_ID = ps.FM_SERV_ID 
                    join MSS_ORDER_SERV os on ps.FM_SERV_ID = os.FM_SERV_ID and os.FM_ORG_ID = ord.FM_ORG_ID
                    join MSS_PROFILES pr on ps.MSS_PROFILES_ID = pr.MSS_PROFILES_ID and pr.FM_ORG_ID = ord.FM_ORG_ID
                    join MSS_PROFILES_BIOMATERIAL prbm on pr.MSS_PROFILES_ID = prbm.MSS_PROFILES_ID
                    join PATIENTS p on pd.PATIENTS_ID = p.PATIENTS_ID 
                    join FM_DEP d on pd.MEDECINS_BIO_DEP_ID = d.FM_DEP_ID
                    join FM_ORG o on d.MAIN_ORG_ID = o.FM_ORG_ID
                    left outer join mss_external_org_codes oe on oe.FM_ORG_ID = o.FM_ORG_ID and oe.EXTERNAL_ORG_ID = ord.FM_ORG_ID
                    where opd.MSS_ORDER_REQUEST_ID = @ord_id
                ", checkinfo.PatientId, Processing.Ini.OrgId).ToList();

            db.Dispose();
        }

        //Получение информации о сформированных заказах
        static public void GetOrderInfoList(out List<Request> RequestInfoList)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);

            Table<OrderRequest> TOrder = db.GetTable<OrderRequest>();
            Table<Patdirec> TPatdirec = db.GetTable<Patdirec>();
            Table<Dep> TDep = db.GetTable<Dep>();
            Table<FM_ORG> TOrg = db.GetTable<FM_ORG>();
            Table<EXT_ORG_CODE> TExtCodes = db.GetTable<EXT_ORG_CODE>();
            RequestInfoList = (from re in TOrder
                               join pe in TPatdirec on re.OrderId equals pe.OrderId
                               join de in TDep on pe.MedBioDepId equals de.DepId
                               join inorg in TOrg on de.MainOrgId equals inorg.OrgId

                               join mainorg in TOrg on inorg.OrgId equals mainorg.OrgId into outerorg
                               from morg in outerorg.DefaultIfEmpty()

                               join incode in TExtCodes on inorg.OrgId equals incode.OrgId into outercodes
                               from code1 in outercodes.DefaultIfEmpty()

                               join maincode in TExtCodes on morg.OrgId equals maincode.OrgId into outermaincodes
                               from code2 in outermaincodes.DefaultIfEmpty()

                               where re.OrgId == Processing.Ini.OrgId && re.State == 0
                               select new Request
                               {
                                   ID = re.OrderId,
                                   Code = re.Code,
                                   Nr = re.OrderNumber,
                                   Date = re.DateOrder,
                                   PatientID = re.PatientId
                               }).Distinct().ToList();
            db.Dispose();
        }

        //Обновление штррих-кода 
        static public void UpdateBarCode(DataContext db, int ord_id)
        {
            DateTime d = DateTime.Now;
            //string code = Message.numberTo32sys(d.Year) + Message.numberTo32sys(d.Month) + Message.numberTo32sys(d.Day) + HospitalCode;
            string code = d.ToString("yyMMdd");

            int success = db.ExecuteCommand(@"
                    update MSS_BARCODE_NUMBERS set BARCODE_NUMBER = 
                                case when Len(b.BARCODE_NUMBER) <= 4 then  {1} + b.BARCODE_NUMBER 
                                     else b.BARCODE_NUMBER end 
                    from mss_order_request r
                    join MSS_ORDER_PATDIREC ppd on r.MSS_ORDER_REQUEST_ID = ppd.MSS_ORDER_REQUEST_ID
                    join PATDIREC p on p.PATDIREC_ID = ppd.PATDIREC_ID
                    join MSS_BARCODE_NUMBERS b on b.PATDIREC_ID = p.PATDIREC_ID
                    where r.MSS_ORDER_REQUEST_ID = {0}
                ", ord_id, code);
        }

        //Получение данных по заказу для формирования XML
        static public void GetOrderInfo(int RequestID, out PatientInfo qPatient, out List<Processing.AssayInfo> AssayInfoList, out List<Processing.Order> OrderList)
        {
            //Создание контекста для работы с базой данных
            DataContext db = new DataContext(Processing.sqlConnectionStr);

            //Данные о пациенте
            qPatient = db.ExecuteQuery<PatientInfo>(@"
                select p.PATIENTS_ID ID, p.NOM LastName, p.PRENOM FirstName, p.PATRONYME MiddleName, p.POL Gender, p.NE_LE BirthDate
                from MSS_ORDER_REQUEST ord
                join PATIENTS p on ord.PATIENTS_ID = p.PATIENTS_ID
                where ord.MSS_ORDER_REQUEST_ID = {0}
            ", RequestID).Single();

            //Данные о пробирках 
            AssayInfoList = db.ExecuteQuery<Processing.AssayInfo>(@"
                select distinct brn.BARCODE_NUMBER Barcode, bm.CODE BiomaterialCode, brn.MSS_BARCODE_NUMBERS_ID BarcodeId
                from MSS_ORDER_REQUEST ord
                join MSS_ORDER_PATDIREC opd on ord.MSS_ORDER_REQUEST_ID = opd.MSS_ORDER_REQUEST_ID
                join PATDIREC pd on opd.PATDIREC_ID = pd.PATDIREC_ID
                join MSS_BARCODE_NUMBERS brn on pd.PATDIREC_ID = brn.PATDIREC_ID and brn.FM_ORG_ID = ord.FM_ORG_ID
                join LAB_BIOTYPE bm on brn.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
                where ord.MSS_ORDER_REQUEST_ID = {0} 
            ", RequestID).ToList();

            //Данные о исследованиях
            OrderList = db.ExecuteQuery<Processing.Order>(@"
                select pr.CODE Code, brn.MSS_BARCODE_NUMBERS_ID BarcodeId, pr.LABEL Name
                from MSS_ORDER_REQUEST ord
                join MSS_ORDER_PATDIREC opd on ord.MSS_ORDER_REQUEST_ID = opd.MSS_ORDER_REQUEST_ID
                join PATDIREC pd on opd.PATDIREC_ID = pd.PATDIREC_ID
                join MSS_BARCODE_NUMBERS brn on pd.PATDIREC_ID = brn.PATDIREC_ID and brn.FM_ORG_ID = ord.FM_ORG_ID
                join LAB_BIOTYPE bm on brn.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
                join DIR_SERV ds on pd.PATDIREC_ID = ds.PATDIREC_ID
                join MSS_PROFILE_SERV ps on ds.FM_SERV_ID = ps.FM_SERV_ID 
                join MSS_ORDER_SERV os on ps.FM_SERV_ID = os.FM_SERV_ID and os.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILES pr on ps.MSS_PROFILES_ID = pr.MSS_PROFILES_ID and pr.FM_ORG_ID = ord.FM_ORG_ID
                left outer join FM_ORG lab on ds.FM_ORG_ID = lab.FM_ORG_ID
                where ord.MSS_ORDER_REQUEST_ID = {0} 
	                and (ISNULL(lab.FM_ORG_ID,0) = 0 or lab.ORG_TYPE = 'I')
            ", RequestID).ToList();
            db.Dispose();
        }

        //Получение данных по заказу для формирования XML
        static public void GetOrderInfoV2(int RequestID, out PatientInfo qPatient, out List<Processing.AssayInfo> AssayInfoList, out List<Processing.Order> OrderList)
        {
            //Создание контекста для работы с базой данных
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            //Данные о пациенте
            qPatient = db.ExecuteQuery<PatientInfo>(@"
                select p.PATIENTS_ID ID, p.NOM LastName, p.PRENOM FirstName, p.PATRONYME MiddleName, p.POL Gender, p.NE_LE BirthDate
                from MSS_ORDER_REQUEST ord
                join PATIENTS p on ord.PATIENTS_ID = p.PATIENTS_ID
                where ord.MSS_ORDER_REQUEST_ID = {0}
            ", RequestID).Single();
             //Данные о пробирках 
            AssayInfoList = db.ExecuteQuery<Processing.AssayInfo>(@"
                select distinct brn.BARCODE_NUMBER Barcode, bm.CODE BiomaterialCode, brn.MSS_BARCODE_NUMBERS_ID BarcodeId
                from MSS_ORDER_REQUEST ord
                join DIR_SERV dsr on ord.MSS_ORDER_REQUEST_ID = dsr.MSS_ORDER_REQUEST_ID
                join PATDIREC pd on dsr.PATDIREC_ID = pd.PATDIREC_ID
                join MSS_BARCODE_NUMBERS brn on pd.PATDIREC_ID = brn.PATDIREC_ID 
								and brn.FM_ORG_ID = ord.FM_ORG_ID
                join LAB_BIOTYPE bm on brn.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
                where ord.MSS_ORDER_REQUEST_ID = {0} 
            ", RequestID).ToList();
            //Данные о исследованиях
            OrderList = db.ExecuteQuery<Processing.Order>(@"
                select  distinct pr.CODE Code, brn.MSS_BARCODE_NUMBERS_ID BarcodeId
							, pr.LABEL Name, cast(pbm.OBLIGATORY as int) Obligatory 
                from MSS_ORDER_REQUEST ord
                join DIR_SERV dsr on ord.MSS_ORDER_REQUEST_ID = dsr.MSS_ORDER_REQUEST_ID
                join PATDIREC pd on dsr.PATDIREC_ID = pd.PATDIREC_ID
                join MSS_BARCODE_NUMBERS brn on pd.PATDIREC_ID = brn.PATDIREC_ID 
								and brn.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILE_SERV ps on dsr.FM_SERV_ID = ps.FM_SERV_ID 
                join MSS_ORDER_SERV os on ps.FM_SERV_ID = os.FM_SERV_ID 
								and os.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILES pr on ps.MSS_PROFILES_ID = pr.MSS_PROFILES_ID 
								and pr.FM_ORG_ID = ord.FM_ORG_ID
				join MSS_PROFILES_BIOMATERIAL pbm on pr.MSS_PROFILES_ID = pbm.MSS_PROFILES_ID
                join LAB_BIOTYPE bm on brn.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
								and pbm.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
                left outer join FM_ORG lab on dsr.FM_ORG_ID = lab.FM_ORG_ID
                where ord.MSS_ORDER_REQUEST_ID = {0}
	                and (ISNULL(lab.FM_ORG_ID,0) = 0)
            ", RequestID).ToList();
            db.Dispose();
        }

        //Получение данных по заказу для формирования XML триггер по типу Астромеда
        static public void GetOrderInfoV3(int RequestID, out PatientInfo qPatient, out List<Processing.AssayInfo> AssayInfoList, out List<Processing.Order> OrderList)
        {
            //Создание контекста для работы с базой данных
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            //Данные о пациенте
            qPatient = db.ExecuteQuery<PatientInfo>(@"
                select p.PATIENTS_ID ID, p.NOM LastName, p.PRENOM FirstName, p.PATRONYME MiddleName, p.POL Gender, p.NE_LE BirthDate
                from MSS_ORDER_REQUEST ord
                join PATIENTS p on ord.PATIENTS_ID = p.PATIENTS_ID
                where ord.MSS_ORDER_REQUEST_ID = {0}
            ", RequestID).Single();
            //Данные о пробирках 
            AssayInfoList = db.ExecuteQuery<Processing.AssayInfo>(@"
                select distinct brn.BARCODE_NUMBER Barcode, bm.CODE BiomaterialCode, brn.MSS_BARCODE_NUMBERS_ID BarcodeId
                                , bm.Label BiomaterialLabel
                from MSS_ORDER_REQUEST ord
                join DIR_SERV dsr on ord.MSS_ORDER_REQUEST_ID = dsr.MSS_ORDER_REQUEST_ID
                join MSS_BARCODE_NUMBERS brn on dsr.MSS_BARCODE_NUMBERS_ID = brn.MSS_BARCODE_NUMBERS_ID
								and brn.FM_ORG_ID = ord.FM_ORG_ID
                join LAB_BIOTYPE bm on brn.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
                where ord.MSS_ORDER_REQUEST_ID = {0} 
            ", RequestID).ToList();
            //Данные о исследованиях
            OrderList = db.ExecuteQuery<Processing.Order>(@"
                select  distinct pr.CODE Code, brn.MSS_BARCODE_NUMBERS_ID BarcodeId, pr.LABEL Name
							, cast(isnull(pbm.OBLIGATORY,0) as int) Obligatory, pr.EXTERNAL_ID ExternalId
                from MSS_ORDER_REQUEST ord
                join DIR_SERV dsr on ord.MSS_ORDER_REQUEST_ID = dsr.MSS_ORDER_REQUEST_ID
                join MSS_BARCODE_NUMBERS brn on dsr.MSS_BARCODE_NUMBERS_ID = brn.MSS_BARCODE_NUMBERS_ID
								and brn.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILE_SERV ps on dsr.FM_SERV_ID = ps.FM_SERV_ID 
                join MSS_ORDER_SERV os on ps.FM_SERV_ID = os.FM_SERV_ID 
								and os.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILES pr on ps.MSS_PROFILES_ID = pr.MSS_PROFILES_ID 
								and pr.FM_ORG_ID = ord.FM_ORG_ID
				join MSS_PROFILES_BIOMATERIAL pbm on pr.MSS_PROFILES_ID = pbm.MSS_PROFILES_ID
                join LAB_BIOTYPE bm on brn.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
								and pbm.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
                left outer join FM_ORG lab on dsr.FM_ORG_ID = lab.FM_ORG_ID
                where ord.MSS_ORDER_REQUEST_ID = {0}
	                --and (ISNULL(lab.FM_ORG_ID,0) = 0)
            ", RequestID).ToList();
            db.Dispose();
        }

        //Получение данных по заказу для формирования XML c параметрами из таблицы DATA_GYNAECOL_STATUS (Гинекологический статус) 
        static public void GetOrderInfoV4(int RequestID, out PatientInfo qPatient, out List<Processing.AssayInfo> AssayInfoList, out List<Processing.Order> OrderList)
        {
            //Создание контекста для работы с базой данных
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            //Данные о пациенте
            qPatient = db.ExecuteQuery<PatientInfo>(@"
               select top 1 p.PATIENTS_ID ID, p.NOM LastName, p.PRENOM FirstName, p.PATRONYME MiddleName
               , p.POL Gender, isnull(p.NE_LE,'1900-01-01') BirthDate
               , isnull(gs.FAZA_CIKLA,0) FazaCikla, cast (isnull(SROK_BEREMENNOSTI_V_NEDEL,0) as int) SrokBeremennosti
               , isnull(ISNULL(E_MAIL, EMAIL),'') Email
                from MSS_ORDER_REQUEST ord
                join PATIENTS p on ord.PATIENTS_ID = p.PATIENTS_ID
                join DIR_SERV dsr on ord.MSS_ORDER_REQUEST_ID = dsr.MSS_ORDER_REQUEST_ID
                join PATDIREC pd on dsr.PATDIREC_ID = pd.PATDIREC_ID
                left outer join MOTCONSU m on pd.MOTCONSU_ID = m.MOTCONSU_ID
                left outer join DATA_GYNAECOL_STATUS gs on gs.MOTCONSU_ID = m.MOTCONSU_ID
                where ord.MSS_ORDER_REQUEST_ID = {0}
                group by p.PATIENTS_ID, p.NOM, p.PRENOM, p.PATRONYME, p.POL, isnull(p.NE_LE,'1900-01-01')
                , gs.FAZA_CIKLA, SROK_BEREMENNOSTI_V_NEDEL, isnull(ISNULL(E_MAIL, EMAIL),'')         
            ", RequestID).Single();
            //Данные о пробирках 
            AssayInfoList = db.ExecuteQuery<Processing.AssayInfo>(@"
                select distinct brc.BARCODE_NUMBER Barcode, bm.CODE BiomaterialCode, brc.MSS_BARCODE_NUMBERS_ID BarcodeId
                                , bm.Label BiomaterialLabel, isnull(bm.TYPE,0) BioType, pr.CODE ProfileCode
                from MSS_ORDER_REQUEST ord
                join DIR_SERV dsr on ord.MSS_ORDER_REQUEST_ID = dsr.MSS_ORDER_REQUEST_ID
				join MSS_BARCODE_SERV brsr on dsr.DIR_SERV_ID = brsr.DIR_SERV_ID
				join MSS_BARCODE_NUMBERS brc on brc.MSS_BARCODE_NUMBERS_ID = brsr.MSS_BARCODE_NUMBERS_ID
								and brc.FM_ORG_ID = ord.FM_ORG_ID
                join LAB_BIOTYPE bm on brc.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
                join MSS_PROFILE_SERV ps on dsr.FM_SERV_ID = ps.FM_SERV_ID 
                join MSS_ORDER_SERV os on ps.FM_SERV_ID = os.FM_SERV_ID 
								and os.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILES pr on ps.MSS_PROFILES_ID = pr.MSS_PROFILES_ID 
								and pr.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILES_BIOMATERIAL prbm on pr.MSS_PROFILES_ID = prbm.MSS_PROFILES_ID
								and prbm.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
                where ord.MSS_ORDER_REQUEST_ID = {0} 
            ", RequestID).ToList();
            
            //Данные о исследованиях
            OrderList = db.ExecuteQuery<Processing.Order>(@"
                select  distinct pr.CODE Code, brc.MSS_BARCODE_NUMBERS_ID BarcodeId, pr.LABEL Name
							, cast(isnull(pbm.OBLIGATORY,0) as int) Obligatory, pr.EXTERNAL_ID ExternalId
                            , isnull(bm.TYPE,0) BioType, bm.CODE BioCode, isnull(OPTIONAL_GROUP,0) GroupOption
                from MSS_ORDER_REQUEST ord
                join DIR_SERV dsr on ord.MSS_ORDER_REQUEST_ID = dsr.MSS_ORDER_REQUEST_ID
				join MSS_BARCODE_SERV brsr on dsr.DIR_SERV_ID = brsr.DIR_SERV_ID
				join MSS_BARCODE_NUMBERS brc on brc.MSS_BARCODE_NUMBERS_ID = brsr.MSS_BARCODE_NUMBERS_ID
								and brc.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILE_SERV ps on dsr.FM_SERV_ID = ps.FM_SERV_ID 
                join MSS_ORDER_SERV os on ps.FM_SERV_ID = os.FM_SERV_ID 
								and os.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILES pr on ps.MSS_PROFILES_ID = pr.MSS_PROFILES_ID 
								and pr.FM_ORG_ID = ord.FM_ORG_ID
				join MSS_PROFILES_BIOMATERIAL pbm on pr.MSS_PROFILES_ID = pbm.MSS_PROFILES_ID
                join LAB_BIOTYPE bm on brc.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
								and pbm.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
                left outer join FM_ORG lab on dsr.FM_ORG_ID = lab.FM_ORG_ID
                 where ord.MSS_ORDER_REQUEST_ID = {0}
	                and (ISNULL(lab.FM_ORG_ID,0) = 0)               
            ", RequestID).ToList();
            db.Dispose();
        }
        //Получение данных по заказу для формирования XML c параметрами из таблицы DATA286 (Гинекологический статус) 
        static public void GetOrderInfoV4_G2(int RequestID, out PatientInfo qPatient, out List<Processing.AssayInfo> AssayInfoList, out List<Processing.Order> OrderList)
        {
            //Создание контекста для работы с базой данных
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            //Данные о пациенте
            qPatient = db.ExecuteQuery<PatientInfo>(@"
               select top 1 p.PATIENTS_ID ID, p.NOM LastName, p.PRENOM FirstName, p.PATRONYME MiddleName
               , p.POL Gender, isnull(p.NE_LE,'1900-01-01') BirthDate
               , isnull(gs.FAZA_CIKLA,0) FazaCikla, cast (isnull(SROK_BEREMENNOSTI_V_NEDEL,0) as int) SrokBeremennosti
                from MSS_ORDER_REQUEST ord
                join PATIENTS p on ord.PATIENTS_ID = p.PATIENTS_ID
                join DIR_SERV dsr on ord.MSS_ORDER_REQUEST_ID = dsr.MSS_ORDER_REQUEST_ID
                join PATDIREC pd on dsr.PATDIREC_ID = pd.PATDIREC_ID
                left outer join MOTCONSU m on pd.MOTCONSU_ID = m.MOTCONSU_ID
                left outer join DATA286 gs on gs.MOTCONSU_ID = m.MOTCONSU_ID
                where ord.MSS_ORDER_REQUEST_ID = {0}
                group by p.PATIENTS_ID, p.NOM, p.PRENOM, p.PATRONYME, p.POL, p.NE_LE, gs.FAZA_CIKLA, SROK_BEREMENNOSTI_V_NEDEL
         
            ", RequestID).Single();
            //Данные о пробирках 
            AssayInfoList = db.ExecuteQuery<Processing.AssayInfo>(@"
                select distinct brc.BARCODE_NUMBER Barcode, bm.CODE BiomaterialCode, brc.MSS_BARCODE_NUMBERS_ID BarcodeId
                                , bm.Label BiomaterialLabel, isnull(bm.TYPE,0) BioType, pr.CODE ProfileCode
                from MSS_ORDER_REQUEST ord
                join DIR_SERV dsr on ord.MSS_ORDER_REQUEST_ID = dsr.MSS_ORDER_REQUEST_ID
				join MSS_BARCODE_SERV brsr on dsr.DIR_SERV_ID = brsr.DIR_SERV_ID
				join MSS_BARCODE_NUMBERS brc on brc.MSS_BARCODE_NUMBERS_ID = brsr.MSS_BARCODE_NUMBERS_ID
								and brc.FM_ORG_ID = ord.FM_ORG_ID
                join LAB_BIOTYPE bm on brc.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
                join MSS_PROFILE_SERV ps on dsr.FM_SERV_ID = ps.FM_SERV_ID 
                join MSS_ORDER_SERV os on ps.FM_SERV_ID = os.FM_SERV_ID 
								and os.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILES pr on ps.MSS_PROFILES_ID = pr.MSS_PROFILES_ID 
								and pr.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILES_BIOMATERIAL prbm on pr.MSS_PROFILES_ID = prbm.MSS_PROFILES_ID
								and prbm.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
                where ord.MSS_ORDER_REQUEST_ID = {0} 
            ", RequestID).ToList();

            //Данные о исследованиях
            OrderList = db.ExecuteQuery<Processing.Order>(@"
                select  distinct pr.CODE Code, brc.MSS_BARCODE_NUMBERS_ID BarcodeId, pr.LABEL Name
							, cast(isnull(pbm.OBLIGATORY,0) as int) Obligatory, pr.EXTERNAL_ID ExternalId
                            , isnull(bm.TYPE,0) BioType, bm.CODE BioCode, isnull(OPTIONAL_GROUP,0) GroupOption
                from MSS_ORDER_REQUEST ord
                join DIR_SERV dsr on ord.MSS_ORDER_REQUEST_ID = dsr.MSS_ORDER_REQUEST_ID
				join MSS_BARCODE_SERV brsr on dsr.DIR_SERV_ID = brsr.DIR_SERV_ID
				join MSS_BARCODE_NUMBERS brc on brc.MSS_BARCODE_NUMBERS_ID = brsr.MSS_BARCODE_NUMBERS_ID
								and brc.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILE_SERV ps on dsr.FM_SERV_ID = ps.FM_SERV_ID 
                join MSS_ORDER_SERV os on ps.FM_SERV_ID = os.FM_SERV_ID 
								and os.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILES pr on ps.MSS_PROFILES_ID = pr.MSS_PROFILES_ID 
								and pr.FM_ORG_ID = ord.FM_ORG_ID
				join MSS_PROFILES_BIOMATERIAL pbm on pr.MSS_PROFILES_ID = pbm.MSS_PROFILES_ID
                join LAB_BIOTYPE bm on brc.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
								and pbm.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
                left outer join FM_ORG lab on dsr.FM_ORG_ID = lab.FM_ORG_ID
                 where ord.MSS_ORDER_REQUEST_ID = {0}
	                and (ISNULL(lab.FM_ORG_ID,0) = 0)               
            ", RequestID).ToList();
            db.Dispose();
        }

        //Получение данных по заказу для формирования XML без таблицы DATA_GYNAECOL_STATUS
        static public void GetOrderInfoV4_1(int RequestID, out PatientInfo qPatient, out List<Processing.AssayInfo> AssayInfoList, out List<Processing.Order> OrderList)
        {
            //Создание контекста для работы с базой данных
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            //Данные о пациенте
            qPatient = db.ExecuteQuery<PatientInfo>(@"
               select top 1 p.PATIENTS_ID ID, p.NOM LastName, p.PRENOM FirstName, p.PATRONYME MiddleName
               , p.POL Gender, isnull(p.NE_LE,'1900-01-01') BirthDate
                from MSS_ORDER_REQUEST ord
                join PATIENTS p on ord.PATIENTS_ID = p.PATIENTS_ID
                join DIR_SERV dsr on ord.MSS_ORDER_REQUEST_ID = dsr.MSS_ORDER_REQUEST_ID
                join PATDIREC pd on dsr.PATDIREC_ID = pd.PATDIREC_ID
                left outer join MOTCONSU m on pd.MOTCONSU_ID = m.MOTCONSU_ID
                where ord.MSS_ORDER_REQUEST_ID = {0}
                group by p.PATIENTS_ID, p.NOM, p.PRENOM, p.PATRONYME, p.POL, isnull(p.NE_LE,'1900-01-01')
         
            ", RequestID).Single();
            //Данные о пробирках 
            AssayInfoList = db.ExecuteQuery<Processing.AssayInfo>(@"
                select distinct brc.BARCODE_NUMBER Barcode, bm.CODE BiomaterialCode, brc.MSS_BARCODE_NUMBERS_ID BarcodeId
                                , bm.Label BiomaterialLabel, isnull(bm.TYPE,0) BioType, pr.CODE ProfileCode
                from MSS_ORDER_REQUEST ord
                join DIR_SERV dsr on ord.MSS_ORDER_REQUEST_ID = dsr.MSS_ORDER_REQUEST_ID
				join MSS_BARCODE_SERV brsr on dsr.DIR_SERV_ID = brsr.DIR_SERV_ID
				join MSS_BARCODE_NUMBERS brc on brc.MSS_BARCODE_NUMBERS_ID = brsr.MSS_BARCODE_NUMBERS_ID
								and brc.FM_ORG_ID = ord.FM_ORG_ID
                join LAB_BIOTYPE bm on brc.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
                join MSS_PROFILE_SERV ps on dsr.FM_SERV_ID = ps.FM_SERV_ID 
                join MSS_ORDER_SERV os on ps.FM_SERV_ID = os.FM_SERV_ID 
								and os.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILES pr on ps.MSS_PROFILES_ID = pr.MSS_PROFILES_ID 
								and pr.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILES_BIOMATERIAL prbm on pr.MSS_PROFILES_ID = prbm.MSS_PROFILES_ID
								and prbm.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
                where ord.MSS_ORDER_REQUEST_ID = {0} 
            ", RequestID).ToList();

            //Данные о исследованиях
            OrderList = db.ExecuteQuery<Processing.Order>(@"
                select  distinct pr.CODE Code, brc.MSS_BARCODE_NUMBERS_ID BarcodeId, pr.LABEL Name
							, cast(isnull(pbm.OBLIGATORY,0) as int) Obligatory, pr.EXTERNAL_ID ExternalId
                            , isnull(bm.TYPE,0) BioType, bm.CODE BioCode, isnull(OPTIONAL_GROUP,0) GroupOption
                from MSS_ORDER_REQUEST ord
                join DIR_SERV dsr on ord.MSS_ORDER_REQUEST_ID = dsr.MSS_ORDER_REQUEST_ID
				join MSS_BARCODE_SERV brsr on dsr.DIR_SERV_ID = brsr.DIR_SERV_ID
				join MSS_BARCODE_NUMBERS brc on brc.MSS_BARCODE_NUMBERS_ID = brsr.MSS_BARCODE_NUMBERS_ID
								and brc.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILE_SERV ps on dsr.FM_SERV_ID = ps.FM_SERV_ID 
                join MSS_ORDER_SERV os on ps.FM_SERV_ID = os.FM_SERV_ID 
								and os.FM_ORG_ID = ord.FM_ORG_ID
                join MSS_PROFILES pr on ps.MSS_PROFILES_ID = pr.MSS_PROFILES_ID 
								and pr.FM_ORG_ID = ord.FM_ORG_ID
				join MSS_PROFILES_BIOMATERIAL pbm on pr.MSS_PROFILES_ID = pbm.MSS_PROFILES_ID
                join LAB_BIOTYPE bm on brc.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
								and pbm.LAB_BIOTYPE_ID = bm.LAB_BIOTYPE_ID
                left outer join FM_ORG lab on dsr.FM_ORG_ID = lab.FM_ORG_ID
                 where ord.MSS_ORDER_REQUEST_ID = {0}
	                and (ISNULL(lab.FM_ORG_ID,0) = 0)               
            ", RequestID).ToList();
            db.Dispose();
        }

        //Получение данных по заказу для формирования XML  для Инвитро
        static public void GetOrderInfoInvitro(int RequestID, out PatientInfo qPatient, out List<OrderInvitro> OrderList)
        {
            //Создание контекста для работы с базой данных
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            //Данные о пациенте
            qPatient = db.ExecuteQuery<PatientInfo>(@"
                select p.PATIENTS_ID ID, p.NOM LastName, p.PRENOM FirstName, p.PATRONYME MiddleName, p.POL Gender, p.NE_LE BirthDate
                        , p.RABOTA workPlace, p.DOLGNOST post, p.MOBIL_TELEFON phone, isnull(p.E_MAIL, p.EMAIL) Email, doctype.LABEL docTypeName
						--, MSS_DOCTYPE.MSS_DOCTYPE_ID docTypeId, MSS_COUNTRY.Name countryName, MSS_COUNTRY.MSS_COUNTRY_ID countryId
                        --, MSS_REGION.MSS_REGION_ID regionId, MSS_CITY.MSS_CITY_ID cityId
						, p.PASPORT_N docNumber, p.SERIQ_DOKUMENTA docSerNumber, p.KOGDA_V_DAN issueDate, p.KEM_V_DAN issueOrgName
						, p.SNILS snils, isnull(oms.SERIQ_POLISA,'') + isnull(oms.NOMER_POLISA,'') oms 
                    	, ADR_OBLAST.NAME regionName, ADR_OBLAST.SOCR regionType, ADR_REGION.NAME district, ADR_REGION.SOCR districtType
                    	, ADR_GOROD.NAME town, ADR_GOROD.SOCR townType,  ADR_STREET.NAME streetName
                    	, ADR_STREET.SOCR streetType, p.DOM house, p.KORPUS building, p.KVARTIRA appartament, p.ADRES_PO_PROPISKE addrReg, p.ADRES_FAKTIHESKIJ addrFact
                from MSS_ORDER_REQUEST ord
                join PATIENTS p on ord.PATIENTS_ID = p.PATIENTS_ID
	             LEFT OUTER JOIN OMI_COUNTRY OMI_COUNTRY ON OMI_COUNTRY.OMS_COUNTRY_ID = p.STRANA1 
                 LEFT OUTER JOIN ADR_OBLAST ADR_OBLAST ON ADR_OBLAST.ADR_OBLAST_ID = p.KOD_TERRITORII 
                 LEFT OUTER JOIN ADR_REGION ADR_REGION ON ADR_REGION.ADR_REGION_ID = p.RAJON_ROSSII 
                 LEFT OUTER JOIN ADR_GOROD ADR_GOROD ON ADR_GOROD.ADR_GOROD_ID = p.OBLAST_GOROD 
                 LEFT OUTER JOIN ADR_STREET ADR_STREET ON ADR_STREET.ADR_STREET_ID = p.ULICA_MOSKVA 
	             --LEFT OUTER JOIN MSS_COUNTRY MSS_COUNTRY ON OMI_COUNTRY.OMS_COUNTRY_ID = MSS_COUNTRY.OMS_COUNTRY_ID
	             --LEFT OUTER JOIN MSS_REGION MSS_REGION ON ADR_OBLAST.ADR_OBLAST_ID = MSS_REGION.ADR_OBLAST_ID
	             --LEFT OUTER JOIN MSS_CITY MSS_CITY ON ADR_GOROD.ADR_GOROD_ID = MSS_CITY.ADR_GOROD_ID
                 LEFT OUTER JOIN OMI_DOCTYPE doctype ON doctype.OMS_DOCTYPE_ID = p.VID_DOKUMENTA
	             --LEFT OUTER JOIN MSS_DOCTYPE MSS_DOCTYPE ON doctype.OMS_DOCTYPE_ID = MSS_DOCTYPE.OMS_DOCTYPE_ID
                 LEFT OUTER JOIN DATA_INS_POLICIES_OMI oms ON (p.PATIENTS_ID = oms.PATIENTS_ID		    
		            and oms.N_LINE = (select MAX(o.N_LINE) from DATA_INS_POLICIES_OMI o where o.PATIENTS_ID = oms.PATIENTS_ID ))  
                 where ord.MSS_ORDER_REQUEST_ID = {0}
            ", RequestID).Single(); 

        //Данные о пробирках 
        OrderList = db.ExecuteQuery<OrderInvitro>(@"
                  select ord.MSS_ORDER_REQUEST_ID RequestId, ord.RequestCode RequestCode,  ord.ORDER_NUMBER RequestNr, ord.DATE_ORDER RequestDate
                           , pd.FM_INTORG_ID FilialId, isnull(oe.Code,'') FilialCode, o.LABEL FilialLabel
                           , pr.EXTERNAL_ID ProductExternalId, prb.OPTION_SET BiomaterialOption, bt.CODE BiomaterialCode
                           , ax.EXTERNAL_ID AuxExtCode, ax.TABLE_NAME AuxTableName,	ax.FIELD_NAME AuxFieldName, pd.MOTCONSU_ID MotconsuId
                    from PATDIREC pd
					join DIR_SERV dsr on dsr.PATDIREC_ID=pd.PATDIREC_ID
					join MSS_ORDER_SERV ors on ors.FM_SERV_ID=dsr.FM_SERV_ID
								and isnull(ors.state, 0) = 1
					join MSS_PROFILE_SERV prs on dsr.FM_SERV_ID = prs.FM_SERV_ID
					join MSS_ORDER_REQUEST ord on dsr.MSS_ORDER_REQUEST_ID = ord.MSS_ORDER_REQUEST_ID
                    join MSS_PROFILES pr on pr.MSS_PROFILES_ID = prs.MSS_PROFILES_ID 
								and pr.FM_ORG_ID = ors.FM_ORG_ID
                    join MSS_PROFILES_BIOMATERIAL prb on pr.MSS_PROFILES_ID = prb.MSS_PROFILES_ID
                    join LAB_BIOTYPE bt on bt.LAB_BIOTYPE_ID = prb.LAB_BIOTYPE_ID 
									and (pd.BIO_TYPE like ('%'+substring(bt.CODE,0,8)+'%') 
										or isnull(pd.BIO_TYPE,'') = '')
                    join FM_DEP d on pd.MEDECINS_BIO_DEP_ID = d.FM_DEP_ID
                    join FM_ORG o on d.MAIN_ORG_ID = o.FM_ORG_ID
                    left outer join mss_external_org_codes oe on oe.FM_ORG_ID = o.FM_ORG_ID 
									and oe.EXTERNAL_ORG_ID = ord.FM_ORG_ID
                    left outer join MSS_PROFILE_AUXILARY pa on pr.MSS_PROFILES_ID = pa.MSS_PROFILES_ID
					left outer join MSS_AUXILARY ax on pa.MSS_AUXILARY_ID = ax.MSS_AUXILARY_ID 
								and ax.IsRequired = 1
                    where ord.MSS_ORDER_REQUEST_ID = {0}
                            and isnull(prb.NOT_USE,0) <> 1
							--and ord.STATE = 0
                    group by ord.MSS_ORDER_REQUEST_ID, ord.RequestCode,  ord.ORDER_NUMBER, ord.DATE_ORDER
                           , pd.FM_INTORG_ID, isnull(oe.Code,''), o.LABEL,  pr.EXTERNAL_ID, prb.OPTION_SET, bt.CODE
                           , ax.EXTERNAL_ID, ax.TABLE_NAME,	ax.FIELD_NAME, pd.MOTCONSU_ID
            ", RequestID).ToList();
            db.Dispose();
        }

        //Создание строк в таблице MSS_BARCODE_NUMBERS
        /*
        static public void InsertBarcodes(int RequestID, Invitro.StickerInfo sticker, int LabId)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int success = db.ExecuteCommand(@"
 						declare @id_t int
						exec up_get_id 'MSS_BARCODE_NUMBERS', 1, @id_t out 
						insert MSS_BARCODE_NUMBERS (MSS_BARCODE_NUMBERS_ID, LAB_BIOTYPE_ID, FM_ORG_ID 
							, BARCODE_NUMBER, STICKER_CODE_BASE64, MSS_ORDER_REQUEST_ID)
						select top 1 @id_t, LAB_BIOTYPE_ID , {4}, {1}, {3}, {0}
						from LAB_BIOTYPE 
						where CODE = {2}
                ", RequestID, sticker.BarCodeNumber, sticker.BiomaterialId, sticker.ZPLText, LabId);
            db.Dispose();
        }
        */
        //Получение доп. иннфрмации для заказа Инвитро
        static public List<Invitro.AuilaryData> GetAuxData(List<Invitro.AuxilaryInfo> auxilaryInfo)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            List<Invitro.AuilaryData> auxilaryDatas = new List<Invitro.AuilaryData>();
            Invitro.AuilaryData auxData;
            foreach (var aux in auxilaryInfo)
            {
                auxData = db.ExecuteQuery<Invitro.AuilaryData>(@"
                    declare @statment nvarchar(max)
					set @statment = N'select case when ISNUMERIC(cast('+ {0} +' as varchar(255)))=1 
					then cast(cast( cast('+ {0} +' as varchar(255)) as float) as varchar(250)) else cast('+ {0} +' as varchar(255)) end 
					AuxValue from '+ {1} +' where MOTCONSU_Id =' + {2}
					EXECUTE sp_executesql @statment                                      
                                    ",  aux.AuxFieldName, aux.AuxTableName, aux.MotconsuId.ToString()).SingleOrDefault();
                
                if (auxData != null)
                {
                    if (auxData.AuxValue != null)
                    {
                        auxData.AuxExtCode = aux.AuxExtCode;
                        auxilaryDatas.Add(auxData);
                    }
                }
            }
            return auxilaryDatas;
        }
        //Получение доп. иннфрмации для заказа
        static public void GetAddInfo(int RequestID, out List<Processing.AddInfo> AddInfoList)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            List<Processing.AddInfo> aInfoList;

            AddInfoList = db.ExecuteQuery<Processing.AddInfo>(@"
                declare @ord_id int = {0}
                select Code ProfileCode, 2 GroupOption, fields InfoId, Value Value
                from (
                  select top 1 Code, [quantity],[diagnoz],[clinic],[instrumental],[treatment],[doctor],[date]
                  from view_add_info_g2 
                  where MSS_ORDER_REQUEST_id = @ord_id and OPTIONAL_GROUP = 2
	                ) as t
                unpivot (
                  Value for fields in ([quantity],[diagnoz],[clinic],[instrumental],[treatment],[doctor],[date])
		                ) as unpvt
            ", RequestID).ToList();

            try
            {
                aInfoList = db.ExecuteQuery<Processing.AddInfo>(@"
                declare @ord_id int = {0}
                select Code ProfileCode, 4 GroupOption, fields InfoId, Value Value
                from (
                  select top 1 Code, [biopsiaType],[fos],[firstBiopsia],[locus],[object],[quantity],[clinic],[instrumental]
		                ,[diagnoz],[treatment],[pregnancy],[cycle],[contraceptive],[menstruation],[cycleRegularity]
		                ,[menopause],[doctor],[date]
                  from view_add_info_g4 
                  where MSS_ORDER_REQUEST_id = @ord_id and OPTIONAL_GROUP = 4
	                ) as t
                unpivot (
                  Value for fields in ([biopsiaType],[fos],[firstBiopsia],[locus],[object],[quantity],[clinic],[instrumental]
		                ,[diagnoz],[treatment],[pregnancy],[cycle],[contraceptive],[menstruation],[cycleRegularity]
		                ,[menopause],[doctor],[date])
		                ) as unpvt
             ", RequestID).ToList();
                foreach (var aInfo in aInfoList) AddInfoList.Add(aInfo);
            }
            catch (Exception ex)
            {
                Files.WriteLogFile("Отсутствует представление: " + ex.Message); Console.WriteLine("Отсутствует представление: " + ex.Message);
            }
            try
            {
                aInfoList = db.ExecuteQuery<Processing.AddInfo>(@" 
                declare @ord_id int = {0}
                select Code ProfileCode, 9 GroupOption, fields InfoId, Value Value
                from (
                  select top 1 Code, [Height],[Weight],[U_Vol] [U-Vol]
                  from view_add_info_g9 
                  where MSS_ORDER_REQUEST_id = @ord_id and OPTIONAL_GROUP = 9
	                ) as t
                unpivot (
                  Value for fields in ([Height],[Weight], [U-Vol])
		                ) as unpvt
            ", RequestID).ToList();
            foreach (var aInfo in aInfoList) AddInfoList.Add(aInfo);
            }
            catch (Exception ex)
            {
                Files.WriteLogFile("Отсутствует представление: " + ex.Message); Console.WriteLine("Отсутствует представление: " + ex.Message);
            }
            try
            {
                aInfoList = db.ExecuteQuery<Processing.AddInfo>(@"
                declare @ord_id int = {0}
                select Code ProfileCode, 8 GroupOption, fields InfoId, Value Value
                from (
                  select top 1 Code, [U_Vol] [U-Vol]
                  from view_add_info_g8 
                  where MSS_ORDER_REQUEST_id = @ord_id and OPTIONAL_GROUP = 8
	                ) as t
                unpivot (
                  Value for fields in ([U-Vol])
		                ) as unpvt
            ", RequestID).ToList();
            foreach (var aInfo in aInfoList) AddInfoList.Add(aInfo);
            }
            catch (Exception ex)
            {
                Files.WriteLogFile("Отсутствует представление: " + ex.Message); Console.WriteLine("Отсутствует представление: " + ex.Message);
            }
            try
            {
                aInfoList = db.ExecuteQuery<Processing.AddInfo>(@"
                declare @ord_id int = {0}
                select Code ProfileCode, 10 GroupOption, fields InfoId, Value Value
                from (
                  select top 1 Code, [PREGNANCY]
                  from view_add_info_g10 
                  where MSS_ORDER_REQUEST_id = @ord_id and OPTIONAL_GROUP = 10
	                ) as t
                unpivot (
                  Value for fields in ([PREGNANCY])
		                ) as unpvt
            ", RequestID).ToList();
            foreach (var aInfo in aInfoList) AddInfoList.Add(aInfo);
            }
            catch (Exception ex)
            {
                Files.WriteLogFile("Отсутствует представление: " + ex.Message); Console.WriteLine("Отсутствует представление: " + ex.Message);
            }
            db.Dispose();
        }
        
        //Получение ID таблицы импорта IMPDATA
        static public int GetImpdataID(DataContext db)
        {
            int ImpdataId = db.ExecuteQuery<int>(@"
                    declare @imp_id int = 0
	                exec up_get_id 'IMPDATA',1, @imp_id out
                    select  @imp_id").Single();
            return ImpdataId;
        }

        //Проверка ID пациента в базе медиалога
        static public bool GetPatientID(DataContext db, string StrPatID, string RecordID, out int PatID)
        {
            Table<Patient> TPatient = db.GetTable<Patient>();
            int p_id = 0;
            int? q_id = null;
            bool ok = false;
            if (int.TryParse(StrPatID, out p_id))
            {
                q_id = (from pe in TPatient
                        where pe.Id == p_id
                        select pe.Id).SingleOrDefault();
                if (q_id == 0)
                {
                    Console.WriteLine("Пациент с ID = {0} не найден в базе медиалога: {1} ", p_id, RecordID);
                    Files.WriteLogFile(string.Format("Пациент с ID = {0} не найден в базе медиалога: {1} ", p_id, RecordID));
                }
                else { Console.WriteLine("Пациент ID = {0}", p_id); ok = true; } //пациент идентифицирован в базе медиалога
            }
            else
            {
                Console.WriteLine("ID пациента не распознан ", RecordID);
                Files.WriteLogFile(string.Format("ID пациента не распознан в {0}", RecordID));
            }

            if (RecordID != null && !ok)
            {
                int pat = db.ExecuteQuery<int>(@"
                        select top 1 PATIENTS_ID 
                        from MSS_ORDER_REQUEST 
                        where RequestCode = {0}
                      ", RecordID).SingleOrDefault();
                if (pat != 0)
                {
                    ok = true; p_id = pat;
                    Console.WriteLine("Пациент ID = {0}", p_id);
                }
            }

            PatID = p_id;
            return ok;
        }

        static public bool GetPatientID(DataContext db, string StrPatID, string RecordID, int OrderID, out int PatID)
        {
            Table<Patient> TPatient = db.GetTable<Patient>();
            int p_id = 0;
            int? q_id = null;
            bool ok = false;
            if (int.TryParse(StrPatID, out p_id))
            {
                q_id = (from pe in TPatient
                        where pe.Id == p_id
                        select pe.Id).SingleOrDefault();
                if (q_id == 0)
                {
                    Console.WriteLine("Пациент с ID = {0} не найден в базе медиалога: {1} ", p_id, RecordID);
                    Files.WriteLogFile(string.Format("Пациент с ID = {0} не найден в базе медиалога: {1} ", p_id, RecordID));
                }
                else { Console.WriteLine("Пациент ID = {0}", p_id); ok = true; } //пациент идентифицирован в базе медиалога
            }
            else
            {
                Console.WriteLine("ID пациента не распознан ", RecordID);
                Files.WriteLogFile(string.Format("ID пациента не распознан в {0}", RecordID));
            }

            if (OrderID != 0 && !ok)
            {
                int pat = db.ExecuteQuery<int>(@"
                        select top 1 PATIENTS_ID 
                        from MSS_ORDER_REQUEST 
                        where MSS_ORDER_REQUEST = {0}
                      ", OrderID).SingleOrDefault();
                if (pat != 0)
                {
                    ok = true; p_id = pat;
                    Console.WriteLine("Пациент ID = {0}", p_id);
                }
            }

            PatID = p_id;
            return ok;
        }
        /*
        static public bool GetPatientID(DataContext db, string RecordID, out int PatID)
        {
            int p_id = 0;
            bool ok = false;
            if (RecordID != null)
            {
                int pat = db.ExecuteQuery<int>(@"
                        select top 1 PATIENTS_ID 
                        from MSS_ORDER_REQUEST 
                        where RequestCode = {0}
                      ", RecordID).SingleOrDefault();
                if (pat != 0)
                {
                    ok = true; p_id = pat;
                    Console.WriteLine("Пациент ID = {0}", p_id);
                }
            }
            PatID = p_id;
            return ok;
        }
        */

        //Получение ID пациента по ФИО и дате рождения
        static public bool GetPatientID(DataContext db, Processing.PatientName Name, Processing.PatientBirthDate BirthDate, string RecordID, out int PatID)
        {
            Table<Patient> TPatient = db.GetTable<Patient>();
            int p_id = 0;
            bool ok = false;
            List<int> p_id_lst = new List<int>();

            p_id_lst = (from pe in TPatient
                        where pe.LastName.Replace('ё', 'е') == Name.LastName.Replace('ё', 'е')
                            && pe.FirstName == Name.FirstName
                            && pe.MiddleName == Name.MiddleName
                            && (pe.BirthDate.Year == BirthDate.BirthYear
                                || BirthDate.BirthYear == 1900)
                        select pe.Id).ToList();

            int p_cnt = p_id_lst.Count;

            if (p_cnt == 0)
            {
                Console.WriteLine("Пациент  {0} {1} {2}, г.р. {3}, не найден в базе медиалога: {4}"
                                           , Name.LastName, Name.FirstName, Name.MiddleName, BirthDate.BirthYear, RecordID);
                Files.WriteLogFile(string.Format("Пациент  {0} {1} {2}, г.р. {3}, не найден в базе медиалога: {4}"
                                           , Name.LastName, Name.FirstName, Name.MiddleName, BirthDate.BirthYear, RecordID));

                List<int> p_id_list = new List<int>();
                if (Name.MiddleName != string.Empty && Name.FirstName != string.Empty)
                {
                    p_id_list = (from pe in TPatient
                                 where pe.LastName.Replace('ё', 'е') == Name.LastName.Replace('ё', 'е')
                                     && (pe.BirthDate.Year == BirthDate.BirthYear || BirthDate.BirthYear == 1900)
                                     &&
                                     ((pe.FirstName.Substring(0, 1) == Name.FirstName.Substring(0, 1) && pe.MiddleName.Substring(0, 1) == Name.MiddleName.Substring(0, 1))
                                     || pe.FirstName == Name.FirstName)
                                     select pe.Id).ToList();
                    int p_count = p_id_list.Count;
                    if (p_count == 0)
                    {
                        Console.WriteLine("Пациент  {0} {1}. {2}., г.р. {3}, не найден в базе медиалога: {4}"
                                                   , Name.LastName, Name.FirstName.Substring(0, 1), Name.MiddleName.Substring(0, 1), BirthDate.BirthYear, RecordID);
                        Files.WriteLogFile(string.Format("Пациент  {0} {1}. {2}., г.р. {3}, не найден в базе медиалога: {4}"
                                                   , Name.LastName, Name.FirstName.Substring(0, 1), Name.MiddleName.Substring(0, 1), BirthDate.BirthYear, RecordID));
                    }
                    if (p_count > 1)
                    {
                        Console.WriteLine("Дублирование пациентов {0} {1}. {2}., г.р. {3} в базе медиалога: {4}"
                                                    , Name.LastName, Name.FirstName.Substring(0, 1), Name.MiddleName.Substring(0, 1), BirthDate.BirthYear, RecordID);
                        Files.WriteLogFile(string.Format("Дублирование пациентов {0} {1}. {2}., г.р. {3} в базе медиалога: {4}"
                                                   , Name.LastName, Name.FirstName.Substring(0, 1), Name.MiddleName.Substring(0, 1), BirthDate.BirthYear, RecordID));
                    }
                    if (p_count == 1)
                    {
                        p_id = p_id_list[0];
                        Console.WriteLine("Пациент идентифицирован по ФИО и г.р. ID = {0}", p_id);
                        ok = true;
                    }

                }

            }
            //Пациент идентифицирован в базе медиалога
            if (p_cnt == 1)
            {
                ok = true;
                p_id = p_id_lst[0];
                Console.WriteLine("Пациент идентифицирован по фамилии, имени, отчеству и г.р., ID = {0}", p_id);
            }
            //Дублирование пациентов
            if (p_cnt > 1)
            {
                p_id_lst = (from pe in TPatient
                            where pe.LastName.Replace('ё', 'е') == Name.LastName.Replace('ё', 'е')
                                && pe.FirstName == Name.FirstName
                                && pe.MiddleName == Name.MiddleName
                                && (pe.BirthDate.Year == BirthDate.BirthYear && pe.BirthDate.Month == BirthDate.BirthMonth && pe.BirthDate.Day == BirthDate.BirthDay)
                            select pe.Id).ToList();
                int p_2cnt = p_id_lst.Count;
                if (p_2cnt == 0)
                {
                    Console.WriteLine("Дублирование пациентов {0} {1}. {2}., г.р. {3} в базе медиалога: {4}"
                                                , Name.LastName, Name.FirstName.Substring(0, 1), Name.MiddleName == string.Empty ? string.Empty : Name.MiddleName.Substring(0, 1), BirthDate.BirthYear, RecordID);
                    Files.WriteLogFile(string.Format("Дублирование пациентов {0} {1}. {2}., г.р. {3} в базе медиалога: {4}"
                                               , Name.LastName, Name.FirstName.Substring(0, 1), Name.MiddleName == string.Empty ? string.Empty : Name.MiddleName.Substring(0, 1), BirthDate.BirthYear, RecordID));
                }
                if (p_2cnt == 1)
                {
                    ok = true;
                    p_id = p_id_lst[0];
                    Console.WriteLine("Пациент идентифицирован по фамилии, имени, отчеству и д.р., ID = {0}", p_id);
                }
                if (p_2cnt > 1)
                {
                    Console.WriteLine("Дублирование пациентов {0} {1}. {2}., д.р. {3} в базе медиалога: {4}"
                                                , Name.LastName, Name.FirstName.Substring(0, 1), Name.MiddleName == string.Empty ? string.Empty : Name.MiddleName.Substring(0, 1), BirthDate, RecordID);
                    Files.WriteLogFile(string.Format("Дублирование пациентов {0} {1}. {2}., д.р. {3} в базе медиалога: {4}"
                                               , Name.LastName, Name.FirstName.Substring(0, 1), Name.MiddleName == string.Empty ? string.Empty : Name.MiddleName.Substring(0, 1), BirthDate, RecordID));
                }
            }

            PatID = p_id;
            return ok;
        }

        //Обновление статуса заказа
        static public void UpdateOrderState(int OrderId, int State, string OrderINZ, string ExternalId)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int success = db.ExecuteCommand(@"
                update mss_order_request set STATE = {1}, RequestCode = {2}, ERROR_STATUS = 'ok', EXTERNAL_ID = {3} 
                where MSS_ORDER_REQUEST_ID = {0}
                ", OrderId, State, OrderINZ, ExternalId);
            db.Dispose();
        }

        //Обновление информации по контейнерам в направлениях заказа 
        static public void UpdatePatdirecContInfo(int OrderId)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int success = db.ExecuteCommand(@"
                    update PATDIREC set COMMENT_BIO = 
	                    stuff((
	                    select char(10) + '---------- ' + 'ИЛН: ' + BARCODE_NUMBER + '  ' + TYPE_CODE + '  ' + TYPE + ' Кол-во: ' + cast(count(BARCODE_NUMBER) as varchar(1)) + 
	                    CHAR(10) + '  ' + 'Тип первичного контейнера: ' + PrimaryType +   
	                    CHAR(10) + '  ' + 'Тип контейнера для ХРАНЕНИЯ и ТРАНСПОРТИРОВКИ: ' + PREPROC +
	                    CHAR(10) + '  ' + 'Исследуемый биоматериал: ' + BiomaterialName + 
	                    CHAR(10) + '  ' + 'Описание контейнера ' + CHAR(10) + cast(DESCR as varchar(2500)) + 
	                    CHAR(10) + '  ' + 'Температура хранения и транспортировки: ' + CONT_TEMP
	                    from MSS_BARCODE_NUMBERS br 
	                    join LAB_CONTTYPE ct on ct.LAB_CONTTYPE_ID = br.LAB_CONTTYPE_ID
	                    where br.MSS_ORDER_REQUEST_ID = drs.MSS_ORDER_REQUEST_ID
	                    group by BARCODE_NUMBER, TYPE_CODE, TYPE,  
	                    PrimaryType, PREPROC, BiomaterialName, cast(DESCR as varchar(2500)), CONT_TEMP
	                    for xml path(''), TYPE).value('.', 'varchar(max)'),1,1,'')
                    from PATDIREC pd
                    join DIR_SERV drs on drs.PATDIREC_ID = pd.PATDIREC_ID
                    where drs.MSS_ORDER_REQUEST_ID = {0}
                ", OrderId);
            db.Dispose();
        }

        static public void UpdateOrderStateError(int OrderId, int State, string ErrorStatus, string ErrorMessage)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int success = db.ExecuteCommand(@"
                update mss_order_request set STATE = {1}, ERROR_STATUS = {2}, ERROR_MESSAGE = {3} 
                where MSS_ORDER_REQUEST_ID = {0}
                ", OrderId, State, ErrorStatus, ErrorMessage);
            db.Dispose();
        }

        static public void UpdateOrderState(int OrderId, int State)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int success = db.ExecuteCommand(@"
                update mss_order_request set STATE = {1}
                where MSS_ORDER_REQUEST_ID = {0}
                ", OrderId, State);
            db.Dispose();
        }

        static public void InsertBarcodes(int OrderId, List<Invitro.StickerInfo> stickersInfo)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var sticker in stickersInfo)
            {
                string[] zplText = sticker.ZPLText.Split(new char[] { (char)10 }, StringSplitOptions.RemoveEmptyEntries);
                
                int cnt = int.TryParse(zplText[zplText.Length - 1].Substring(1, 1), out cnt) ? cnt : 1;
                if (sticker.ContainerId == "e8b803dc-774a-42be-8932-550a13f5f91a") cnt = 1;
                if (sticker.ContainerId == "738fe579-1918-44a7-a35f-4c6be617c02b") cnt = 1;
                
                Console.WriteLine("Количество этикеток {1}: {0}", cnt, sticker.BarCodeNumber);
                for (int i = 1; i <= cnt; i++)
                {
                    int success = db.ExecuteCommand(@"
                    declare @id int
                    exec up_get_id 'MSS_BARCODE_NUMBERS', 1, @id out
                    insert MSS_BARCODE_NUMBERS (MSS_BARCODE_NUMBERS_ID, MSS_ORDER_REQUEST_ID, BARCODE_NUMBER, LAB_BIOTYPE_ID, LAB_CONTTYPE_ID, FM_ORG_ID  )
                    values (@id, {0}, {1}, (select top 1 LAB_BIOTYPE_ID from LAB_BIOTYPE where CODE = {2})
		                    , (select top 1  LAB_CONTTYPE_ID from  LAB_CONTTYPE where EXT_CODE = {3}), {4} )
                ", OrderId, sticker.BarCodeNumber, sticker.BiomaterialId, sticker.ContainerId, Processing.Ini.OrgId);
                }
            }
            db.Dispose();
        }

        static public void UpdateBarcodeProfile(int OrderId, List<Invitro.StickerProfile> stickersProfile)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var profile in stickersProfile)
            {
                int success = db.ExecuteCommand(@"
 
                ", OrderId, profile.ProductId, profile.INZ);

            }
            db.Dispose();
        }

        //Обновление статуса заказа
        static public void UpdateOrderInfo(int OrderId, int State, string Code)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int success = db.ExecuteCommand(@"
                update mss_order_request set STATE = {1}, RequestCode = {2} where MSS_ORDER_REQUEST_ID = {0}
                ", OrderId, State, Code);
            db.Dispose();
        }

        //Раскодировка паролей для входа на вэб сервис
        static public List<Processing.WebFilialInfo> GetWebFilealsInfo()
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            List<Processing.WebFilialInfo> FInfoList = new List<Processing.WebFilialInfo>();
            try
            {
                FInfoList = db.ExecuteQuery<Processing.WebFilialInfo>(@"
                        select CODE Code, WEB_ID WebId, WEB_PSWD Pass 
                        from MSS_EXTERNAL_ORG_CODES
                        where EXTERNAL_ORG_ID = {0} 
                            and Len(WEB_PSWD)>40
                        ", Processing.Ini.OrgId).ToList();

                if (FInfoList.Count == 0)
                {
                    FInfoList.Add(new Processing.WebFilialInfo { Code = Processing.Ini.mis, Pass = Processing.Ini.WPass });
                }
                else
                    foreach (var item in FInfoList)
                    {
                        item.Pass = Crypto.DecryptStringAES(Crypto.DStr(item.Pass), "hip-hop-lab");
                    }
            }
            catch (SqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Настройка филиалов отсутствует!");
                Files.WriteLogFile(string.Format("Настройка филиалов отсутствует! /n {0}", ex.Message));
                FInfoList.Add(new Processing.WebFilialInfo { Code = Processing.Ini.mis, Pass = Processing.Ini.WPass });
            }
            catch (Exception ex)
            {
                Files.WriteLogFile(ex.ToString()); Console.WriteLine(ex.Message);
            }

            return FInfoList;
        }

        static public string GetToken(int FilialId)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            string token;
            token = db.ExecuteQuery<string>(@"
                        select WEB_PSWD Token
                        from MSS_EXTERNAL_ORG_CODES
                        where EXTERNAL_ORG_ID = {0} 
                             and FM_ORG_ID = {1} 
                        ", Processing.Ini.OrgId, FilialId).Single();
            db.Dispose(); 
            return token;
        }

        
        static public FilialInfo GetFilialInfo(int FilialId)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            FilialInfo filialInfo;
            filialInfo = db.ExecuteQuery<FilialInfo>(@"
                        select WEB_PSWD Token, WEB_ID ClientId
                        from MSS_EXTERNAL_ORG_CODES
                        where EXTERNAL_ORG_ID = {0} 
                             and FM_ORG_ID = {1} 
                        ", Processing.Ini.OrgId, FilialId).Single();
            db.Dispose();
            return filialInfo;
        }

        //Проверка изменения контрольной суммы для таблицы MSS_BARCODE_NUMBERS
        static public int? bcChanged()
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int? state = db.ExecuteQuery<int?>(@"
                        declare @state int
						SELECT  top 1 @state = isnull(state,0)
                        FROM MSS_ORDER_REQUEST
						where  DATEDIFF(d,DATE_ORDER,GETDATE()) < 3 and isnull(state,0) in (1,-1)
						SELECT @state
                        ").FirstOrDefault();
            db.Dispose();

            return state;
        }

        static public void UpdateChangedOrder(int state)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int status = db.ExecuteCommand(@"
                    if ({0} = 1)
                        update MSS_ORDER_REQUEST set state = 2 
                        where DATEDIFF(d,DATE_ORDER,GETDATE()) < 3 and isnull(state,0) = 1
                    if ({0} = -1)
                        update MSS_ORDER_REQUEST set state = -2 
                        where DATEDIFF(d,DATE_ORDER,GETDATE()) < 3 and isnull(state,0) = -1
                    ", state);
        }

        //Проверка изменения контрольной суммы для таблицы MSS_BARCODE_NUMBERS
        static public bool delChanged(out int checksum)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            checksum = db.ExecuteQuery<int>(@"
                        SELECT  CHECKSUM_AGG( cast(TO_DELETE as int))
                        FROM MSS_ORDER_REQUEST
                        ").Single();
            db.Dispose();
            bool ch = false;
            //if (Processing.BarCodeCheckSum == 0) Processing.BarCodeCheckSum = checksum;
            ch = (Processing.DeleteCheckSum != checksum);
            if (ch) { Processing.DeleteCheckSum = checksum; Console.WriteLine("Удаление: {0}", checksum); }
            return ch;
        }

        //Заполнение лабораторных справочников по пришедшим методикам
        static public void BuildLabDitionary(DataContext db, int ResGroupId, int DataType)
        {
            int success = db.ExecuteCommand(@"
                    declare @labels table(ID int
					    , LAB_METHODS_ID int
					    , Label varchar(150)
					    , Code varchar(150)
					    )

                    insert @labels
                    select ROW_NUMBER() OVER(order by lm.LAB_METHODS_ID), lm.LAB_METHODS_ID,  lm.LABEL, lm.CODE
                    from LAB_METHODS lm
                    left outer join Lab_methodBio lb on lm.LAB_METHODS_ID = lb.LAB_METHODS_ID
                    where LAB_DEVS_ID = {0}
	                    and LAB_METHODBIO_ID is null
	                    and lm.DATATYPE = {2}
                    order by lm.LAB_METHODS_ID

                    declare @dict_id int
                    set @dict_id = (select top 1 DS_OUTER_DICT_ID 
                                    from ds_outer_dict 
                                    where CURRENTDICT = 1
                                   )

                    declare @i int, @count int
                    declare @out_id int, @p_id int, @lmb_id int, @dic_id int, @rel_id int 
                    set @count = (select COUNT(*) from @labels)
                    set @i=1
                    while @i<=@count
                    begin
                        exec dbo.[up_get_id] 'DS_OUTER_PARAMS', 1, @out_id output
                        insert DS_OUTER_PARAMS (DS_OUTER_PARAMS_ID,LABEL,CODE)
                        select  @out_id, Label, Code
	                    from @labels 
	                    where ID = @i

                        exec dbo.[up_get_id] 'DS_PARAMS', 1, @p_id output
                        insert DS_PARAMS (DS_PARAMS_ID, LABEL, DS_OUTER_PARAMS_ID)
                        select  @p_id, Label, @out_id
	                    from @labels 
	                    where ID = @i
	
                        exec dbo.[up_get_id] 'LAB_METHODBIO', 1, @lmb_id output
                        insert LAB_METHODBIO (LAB_METHODBIO_ID, LAB_METHODS_ID, DS_PARAMS_ID)
                        select  @lmb_id, LAB_METHODS_ID, @p_id
	                    from @labels 
	                    where ID = @i
	
	                    exec dbo.[up_get_id] 'DS_OUTER_DICTPARAMS', 1, @dic_id output
                        insert DS_OUTER_DICTPARAMS (DS_OUTER_DICTPARAMS_ID, LABEL, DS_OUTER_DICT_ID, DS_OUTER_PARAMS_ID, IS_TOPLEVEL, SHOWINRESUME)
                        select  @dic_id, Label, @dict_id, @out_id, 0, 0
	                    from @labels 
	                    where ID = @i                

	                    exec dbo.[up_get_id] 'DS_OUTER_RELATIONS', 1, @rel_id output
	                    insert into DS_OUTER_RELATIONS (DS_OUTER_RELATIONS_ID, PARENT_ID, CHILD_ID, REL_LEVEL, IS_PRIMARY)
	                    select  @rel_id, {1}, @dic_id, 0, 1
	                    from @labels 
	                    where ID = @i 

	                    set @i=@i+1
                    end
                ", Processing.Ini.AnalyzerID, ResGroupId, DataType);
        }
        
        //Заполнение лабораторных справочников 
        static public void BuildDitionary(DataContext db, string Label, string Code, string Num, string Unit, int ResGroupId, out int MethodId)
        {
            MethodId = db.ExecuteQuery<int>(@"
                    declare @label varchar(620), @code varchar(150), @unit varchar(20), @dev_id int, @grp_id int, @num varchar(30)

                        set @label  = {0}
						set @code  = {1}
						set @unit  = {2}
						set @dev_id  = {3}
						set @grp_id  = {4}
                        set @num  = {5}
                     
                     declare @labm_id int, @out_id int, @p_id int, @lmb_id int, @dic_id int, @rel_id int, @dict_id int 
                     set @dict_id = (select top 1 DS_OUTER_DICT_ID 
                                    from ds_outer_dict 
                                    where CURRENTDICT = 1
                                   )
                    
                     exec up_get_id 'LAB_METHODS', 1, @labm_id out
                     insert into LAB_METHODS (LAB_METHODS_ID, LABEL, CODE, UNIT, LAB_DEVS_ID, STATUS, DATATYPE)
                     values (@labm_id, SUBSTRING(@label,0,250), @code, @unit, @dev_id, 1, 0)

                     exec dbo.[up_get_id] 'DS_OUTER_PARAMS', 1, @out_id output
                     insert DS_OUTER_PARAMS (DS_OUTER_PARAMS_ID,LABEL,CODE)
                     values ( @out_id, SUBSTRING(@label,0,150), @Code )

                     exec dbo.[up_get_id] 'DS_PARAMS', 1, @p_id output
                     insert DS_PARAMS (DS_PARAMS_ID, LABEL, DS_OUTER_PARAMS_ID, GRP_ORD)
                     values (@p_id, SUBSTRING(@label,0,150), @out_id, @p_id)
	
                     exec dbo.[up_get_id] 'LAB_METHODBIO', 1, @lmb_id output
                     insert LAB_METHODBIO (LAB_METHODBIO_ID, LAB_METHODS_ID, DS_PARAMS_ID)
                     values (@lmb_id, @labm_id, @p_id)

	
	                 exec dbo.[up_get_id] 'DS_OUTER_DICTPARAMS', 1, @dic_id output
                     insert DS_OUTER_DICTPARAMS (DS_OUTER_DICTPARAMS_ID, LABEL, DS_OUTER_DICT_ID, DS_OUTER_PARAMS_ID, IS_TOPLEVEL, SHOWINRESUME, NUM)
                     values (@dic_id, @label, @dict_id, @out_id, 0, 0, @num)

	                 exec dbo.[up_get_id] 'DS_OUTER_RELATIONS', 1, @rel_id output
	                 insert into DS_OUTER_RELATIONS (DS_OUTER_RELATIONS_ID, PARENT_ID, CHILD_ID, REL_LEVEL, IS_PRIMARY)
	                 values (@rel_id, @grp_id, @dic_id, 0, 1)
                     
                     select @labm_id
                ", Label, Code, Unit, Processing.Ini.AnalyzerID, ResGroupId, Num).Single();        
        }

        static public void BuildMethods(DataContext db, string Label, string Code, string Num, string Unit, out int MethodId)
        {
            MethodId = db.ExecuteQuery<int>(@"
                    declare @label varchar(620), @code varchar(150), @unit varchar(20), @dev_id int

                        set @label  = {0}
						set @code  = {1}
						set @unit  = {2}
						set @dev_id  = {3}

                     declare @labm_id int
                    
                     exec up_get_id 'LAB_METHODS', 1, @labm_id out
                     insert into LAB_METHODS (LAB_METHODS_ID, LABEL, CODE, UNIT, LAB_DEVS_ID, STATUS, DATATYPE)
                     values (@labm_id, SUBSTRING(@label,0,250), @code, @unit, @dev_id, 1, 0)
                     
                     select @labm_id
                ", Label, Code, Unit, Processing.Ini.AnalyzerID).SingleOrDefault();
        }

        static public void BuildDitionary(DataContext db, int MethodId, string Num, string Unit, int ResGroupId)
        {
            MethodId = db.ExecuteQuery<int>(@"
                    declare @labm_id  int, @unit varchar(20), @dev_id int, @grp_id int, @num varchar(30)

                        set @labm_id = {0}
						set @unit = {1}
						set @dev_id = {2}
						set @grp_id = {3}
                        set @num = {4}
                     
                     declare @out_id int, @p_id int, @lmb_id int, @dic_id int, @rel_id int, @dict_id int, @label varchar(250),  @code varchar(100)
                     set @dict_id = (select top 1 DS_OUTER_DICT_ID 
                                    from ds_outer_dict 
                                    where CURRENTDICT = 1
                                   )
                     
                    select @label = LABEL, @code = CODE from LAB_METHODS
                    where LAB_METHODS_ID = @labm_id

                     exec dbo.[up_get_id] 'DS_OUTER_PARAMS', 1, @out_id output
                     insert DS_OUTER_PARAMS (DS_OUTER_PARAMS_ID,LABEL,CODE)
                     values ( @out_id, SUBSTRING(@label,0,150), @Code )

                     exec dbo.[up_get_id] 'DS_PARAMS', 1, @p_id output
                     insert DS_PARAMS (DS_PARAMS_ID, LABEL, DS_OUTER_PARAMS_ID, GRP_ORD)
                     values (@p_id, SUBSTRING(@label,0,150), @out_id, @p_id)
	
                     exec dbo.[up_get_id] 'LAB_METHODBIO', 1, @lmb_id output
                     insert LAB_METHODBIO (LAB_METHODBIO_ID, LAB_METHODS_ID, DS_PARAMS_ID)
                     values (@lmb_id, @labm_id, @p_id)

	
	                 exec dbo.[up_get_id] 'DS_OUTER_DICTPARAMS', 1, @dic_id output
                     insert DS_OUTER_DICTPARAMS (DS_OUTER_DICTPARAMS_ID, LABEL, DS_OUTER_DICT_ID, DS_OUTER_PARAMS_ID, IS_TOPLEVEL, SHOWINRESUME, NUM)
                     values (@dic_id, @label, @dict_id, @out_id, 0, 0, @num)

	                 exec dbo.[up_get_id] 'DS_OUTER_RELATIONS', 1, @rel_id output
	                 insert into DS_OUTER_RELATIONS (DS_OUTER_RELATIONS_ID, PARENT_ID, CHILD_ID, REL_LEVEL, IS_PRIMARY)
	                 values (@rel_id, @grp_id, @dic_id, 0, 1)
                     
                     select @labm_id
                ", MethodId, Unit, Processing.Ini.AnalyzerID, ResGroupId, Num).Single();
        }

        //Создание вида исследования 
        static public void CreateResGroup(DataContext db, string Label, string Num, int ResGroupId, out int ResGroupNewId)
        {
            ResGroupNewId = db.ExecuteQuery<int>(@"
                     declare @label varchar(620) = {0}
						, @num varchar(30) = {1}
						, @grp_id int = {2}
                     
                     declare @dic_id int, @rel_id int, @dict_id int, @parent_id int
	                 
	                 set @dict_id = (select top 1 DS_OUTER_DICT_ID 
                     from ds_outer_dict 
                     where CURRENTDICT = 1
                             )
                     set @parent_id = (select top 1 PARENT_ID 
					 from DS_OUTER_RELATIONS 
					 where CHILD_ID = @grp_id
							)
	                 exec dbo.[up_get_id] 'DS_OUTER_DICTPARAMS', 1, @dic_id output
                     insert DS_OUTER_DICTPARAMS (DS_OUTER_DICTPARAMS_ID, LABEL, DS_OUTER_DICT_ID, NUM, IS_TOPLEVEL, SHOWINRESUME)
                     values (@dic_id, @label, @dict_id, @num, 1, 1)

	                 exec dbo.[up_get_id] 'DS_OUTER_RELATIONS', 1, @rel_id output
	                 insert into DS_OUTER_RELATIONS (DS_OUTER_RELATIONS_ID, PARENT_ID, CHILD_ID, REL_LEVEL, IS_PRIMARY)
	                 values (@rel_id, @parent_id, @dic_id, 0, 1)
                     select @dic_id
                ", Label, Num, ResGroupId).Single();
        }

        static public void CreateResGroupV2(DataContext db, string Label, string Num, int ResGroupId, out int ResGroupNewId)
        {
            ResGroupNewId = db.ExecuteQuery<int>(@"
                     declare @label varchar(620), @num varchar(30), @grp_id int;
                     
                     set @label = {0}
					 set @num  = {1}
					 set @grp_id = {2}

                     declare @dic_id int, @rel_id int, @dict_id int, @parent_id int
	                 
	                 set @dict_id = (select top 1 DS_OUTER_DICT_ID 
                     from ds_outer_dict 
                     where CURRENTDICT = 1
                             )
                     set @parent_id = @grp_id
							
	                 exec dbo.[up_get_id] 'DS_OUTER_DICTPARAMS', 1, @dic_id output
                     insert DS_OUTER_DICTPARAMS (DS_OUTER_DICTPARAMS_ID, LABEL, DS_OUTER_DICT_ID, NUM, IS_TOPLEVEL, SHOWINRESUME)
                     values (@dic_id, @label, @dict_id, @num, 1, 1)

	                 exec dbo.[up_get_id] 'DS_OUTER_RELATIONS', 1, @rel_id output
	                 insert into DS_OUTER_RELATIONS (DS_OUTER_RELATIONS_ID, PARENT_ID, CHILD_ID, REL_LEVEL, IS_PRIMARY)
	                 values (@rel_id, @parent_id, @dic_id, 0, 1)
                     select @dic_id
                ", Label, Num, ResGroupId).Single();
        }

        //Проверка настройки методик
        static public int LabMethodCheck(DataContext db, string MethodCode)
        {
            int labMethodId = db.ExecuteQuery<int>(@"
                    select LAB_METHODS_ID 
                    from LAB_METHODS
                    where LAB_DEVS_ID = {0}
	                    and CODE = {1}
                ", Processing.Ini.AnalyzerID, MethodCode).SingleOrDefault();

            return labMethodId;
        }

        //Проверка настройки методик
        static public int LabMethodCheck(DataContext db, string MethodCode, out bool InStruture)
        {
            int labMethodId = db.ExecuteQuery<int>(@"
                    select LAB_METHODS_ID 
                    from LAB_METHODS
                    where LAB_DEVS_ID = {0}
	                    and CODE = {1}
                ", Processing.Ini.AnalyzerID, MethodCode).SingleOrDefault();
            InStruture = false;
            if (labMethodId != 0)
            {
                int structId = db.ExecuteQuery<int>(@"
                    select lm.LAB_METHODS_ID 
                    from LAB_METHODS lm
                    left outer join LAB_METHODBIO lb on lm.LAB_METHODS_ID = lb.LAB_METHODS_ID
                    where lm.LAB_DEVS_ID = {0}
	                    and lm.LAB_METHODS_ID = {1}
                        and lb.LAB_METHODBIO_ID is not null
                ", Processing.Ini.AnalyzerID, labMethodId).SingleOrDefault();
                if (structId == 0) InStruture = false;
                else InStruture = true;
            }
            return labMethodId;
        }

        //Проверка настройки видов исследований
        static public int ResGroupCheck(DataContext db, string Num, int ResGrpMainId)
        {
            int ResGrpId = db.ExecuteQuery<int>(@"
                    select d.DS_OUTER_DICTPARAMS_ID from DS_OUTER_DICTPARAMS d
                    join DS_OUTER_RELATIONS rp on d.DS_OUTER_DICTPARAMS_ID = rp.CHILD_ID
                    join DS_OUTER_RELATIONS rc on rp.PARENT_ID = rc.PARENT_ID
                    where  rc.CHILD_ID = {0} and NUM = {1}
                          and d.DS_OUTER_PARAMS_ID is null
                ", ResGrpMainId, Num).SingleOrDefault();
            return ResGrpId;
        }

        static public int ResGroupCheckV2(DataContext db, string Num, int ResGrpMainId)
        {
            int ResGrpId = db.ExecuteQuery<int>(@"
                  select d.DS_OUTER_DICTPARAMS_ID from DS_OUTER_DICTPARAMS d
                    join DS_OUTER_RELATIONS rp on d.DS_OUTER_DICTPARAMS_ID = rp.CHILD_ID
                    where  rp.PARENT_ID = {0} and NUM = {1}
                          and d.DS_OUTER_PARAMS_ID is null
                ", ResGrpMainId, Num).SingleOrDefault();
            return ResGrpId;
        }


        //Вставка длинных комментариев к параметрам
        static public void ParamsComments(int ImpdataId, int motconsuId, int patientId, string TableName, string FieldName)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int texist = db.ExecuteQuery<int>(@"
                    select count(*)
                    from METAFIELD 
                    where TABLE_NAME = {0} 
	                    and FIELD_NAME = {1} 
                    ", TableName ?? "-", FieldName ?? "-").SingleOrDefault();
            if (texist != 0)
            {
                string ParamComments = db.ExecuteQuery<string>(@"
						select isnull((select distinct char(10) + 
                            UPPER(dic.LABEL) + 
		                        case when len(isnull(VAL,'')) > 60 then  ' : ' + VAL 
		                            else '' end + 
                                case when DataLength(isnull(M_VAL,'')) > 60 then + ' : ' + cast(M_VAL as varchar(max))
		                            else '' end + 
		                        case when DataLength(isnull(COMMENT,'')) > 60 then ' : ' + VAL + ' ' + isnull(dr.UNIT,'') + char(10) + char(10) + '   Комментарий: ' + char(10) + cast(COMMENT as varchar(max))
		                            else '' end
                            from DS_RESTESTS dr
                            join LAB_METHODBIO lb on dr.LAB_METHODS_ID = lb.LAB_METHODS_ID
                            join DS_PARAMS p on lb.DS_PARAMS_ID = p.DS_PARAMS_ID
                            join VIEW_GRPPRM dic on p.DS_PARAMS_ID = dic.DS_PARAMS_ID
                            join LAB_METHODS lm on lm.LAB_METHODS_ID = dr.LAB_METHODS_ID
                            where ( len(VAL) > 50 or DataLength(COMMENT) > 50 or DataLength(M_VAL) >50 )
                                and lm.CODE not in ('ReqComment', 'AnaComment', 'AnaName', 'Title1', 'Комментарии') 
                                and dic.SHOWINRESUME = 1
                                and dr.IMPDATA_ID = {0}
								and (DataLength(isnull(M_VAL,'')) > 60 or len(isnull(VAL,'')) > 60 or DataLength(isnull(COMMENT,'')) > 60)
						for xml path(''), TYPE).value('.', 'varchar(max)'),'')
                        ", ImpdataId).Single();
                if (ParamComments.Length > 50)
                {
                    ParamComments = ParamComments.Substring(1).Replace("&lt;", "<").Replace("&gt;", ">");

                    int state = db.ExecuteCommand(@"
                            declare @insert_string nvarchar(max)
                            declare @table_name nvarchar(30)
                            declare @field_name nvarchar(30)
				            set @table_name = {3}
                            set @field_name  = {4}
				            set @insert_string=
					            N'if not exists(select top 1 MOTCONSU_ID from '+@table_name+' where MOTCONSU_ID=@m_id) 
					            begin
						            declare @t_id int
						            exec up_get_id '+@table_name+',1,@t_id out
						            insert into '+@table_name+' ('+@table_name+'_ID, PATIENTS_ID,DATE_CONSULTATION, MOTCONSU_ID, '+@field_name+') 
						            values (@t_id, @patients_id, Getdate(), @m_id, @comment)
					             end
					             else update '+@table_name+' set '+@field_name+'=@comment where MOTCONSU_ID=@m_id'
				            exec sp_executesql @insert_string, 
				            N'@patients_id int, @m_id int, @comment varchar(max)',
				            {1}, {0}, {2}
                        ", motconsuId, patientId, ParamComments, TableName, FieldName);
                }
            }
            else
            {
                Console.WriteLine("Ошибка в названиях таблицы {0} и/или поля {1} для вставки комментраиев", TableName, FieldName);
                Files.WriteLogFile(string.Format("Ошибка в названиях таблицы {0} и/или поля {1} для вставки комментраиев", TableName, FieldName));
            }
            db.Dispose();
        }

        public static void InsertBiomaterial(List<Biomat> Biomaterial, int OrgId)
        {
            Console.WriteLine("Вставка биоматериала");
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var b in Biomaterial)
            {
                int success = db.ExecuteCommand(@"
                    declare @id int
                    if not exists (select code from LAB_BIOTYPE where CODE = {0} and FM_ORG_ID = {2})
                    begin
	                    exec up_get_id 'LAB_BIOTYPE',1,@id out
                        insert LAB_BIOTYPE(LAB_BIOTYPE_ID,  CODE, LABEL, FM_ORG_ID)
                        values (@id, {0}, {1}, {2} )
                     end
                    ", b.Code, b.Label, OrgId);
            }
            db.Dispose(); 
        }

        public static void InsertContainer(List<Invitro.TestTube> testTubes, int OrgId)
        {
            Console.WriteLine("Вставка контейнера");
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var t in testTubes)
            {
                int success = db.ExecuteCommand(@"
                    declare @id int
                    if not exists (select code from LAB_CONTTYPE where CODE = {0} )
                    begin
	                    exec up_get_id 'LAB_CONTTYPE',1,@id out
                        insert LAB_CONTTYPE (LAB_CONTTYPE_ID, CODE, LABEL, DESCR, EXT_CODE, TYPE, TYPE_CODE)
                        values (@id, {0}, {1}, {2}, {3}, {4}, {5} )
                     end
                    ", t.Code, t.Name, t.Description, t.ExternalId, t.Type, t.TypeCode);
            }
            db.Dispose();
        }

        public static void InsertProfiles(List<Profile> Profiles, int OrgId)
        {
            Console.WriteLine("Вставка профилей");
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var profile in Profiles)
            {
                int success = db.ExecuteCommand(@"
                    declare @id int
                    if not exists (select code from MSS_PROFILES where EXTERNAL_ID = {3} and FM_ORG_ID = {2} )
                    begin
	                    exec up_get_id 'MSS_PROFILES',1,@id out
                        insert MSS_PROFILES(MSS_PROFILES_ID,  CODE, LABEL, FM_ORG_ID, EXTERNAL_ID, SHORT_LABEL, BLANK, SubgroupName )
                        values (@id, {0}, {1}, {2}, case when {3} <> '' then {3} else null end,
                            case when {4} <> '' then {4} else null end, {5}, {6} )
                     end
                    ", profile.Code, profile.Label, OrgId, profile.ExternalId, profile.ShortLabel, profile.GroupName, profile.SubgroupName);
            }
            db.Dispose();
        }

        public static void InsertExtendedProfiles(List<Invitro.ExtendedProduct> Profiles)
        {
            Console.WriteLine("Вставка расширения профилей");
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var profile in Profiles)
            {
                //Console.WriteLine(profile.ExternalId);
                int success = db.ExecuteCommand(@"
                    declare @id int
                    if not exists (select EXTERNAL_ID from MSS_EXTENDED_PROFILES where EXTERNAL_ID = {0}  
                                        and OPTIONSET = {1} and BIOMATERIAL_ID = {2} and CONTAINER_ID = {3})
                    begin
	                    exec up_get_id 'MSS_EXTENDED_PROFILES',1,@id out
                        insert MSS_EXTENDED_PROFILES(MSS_EXTENDED_PROFILES_ID, EXTERNAL_ID, OPTIONSET, BIOMATERIAL_ID, CONTAINER_ID, FORM_TYPE  )
                        values (@id, {0}, {1}, {2}, {3}, {4} )
                     end
                    ", profile.ExternalId, profile.OptionSet, profile.BiomaterialId, profile.ContainerId, profile.FormType);
            }
            db.Dispose();
        }

        public static void InsertBiomatProfile(List<ProfileBiomat> ProfileBiomats, int OrgId)
        {
            Console.WriteLine("Вставка связи биоматериалов с профилем");
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var prb in ProfileBiomats)
            {
                /*Console.WriteLine("Тест 1");*/
                int success = db.ExecuteCommand(@"
                        declare @id int, @profile_id int, @biomat_id int
                        if not exists (select pb.MSS_PROFILES_BIOMATERIAL_ID 
                            from MSS_PROFILES_BIOMATERIAL pb
                            join MSS_PROFILES p on pb.MSS_PROFILES_ID = p.MSS_PROFILES_ID
                            join LAB_BIOTYPE b on pb.LAB_BIOTYPE_ID = b.LAB_BIOTYPE_ID
                            where b.CODE = {0} and p.EXTERNAL_ID = {4} and p.FM_ORG_ID = {2} and b.FM_ORG_ID = {2} and isnull(pb.OPTION_SET,'') = isnull({3},'')
                                       )
                            begin
                                set @profile_id = (select top 1 MSS_PROFILES_ID from MSS_PROFILES where CODE = {1} and FM_ORG_ID = {2})
                                set @biomat_id = (select top 1 LAB_BIOTYPE_ID from LAB_BIOTYPE where CODE = {0} and FM_ORG_ID = {2})
                                if @profile_id is not null and @biomat_id is not null
                                begin
	                                exec up_get_id 'MSS_PROFILES_BIOMATERIAL',1,@id out
                                    insert MSS_PROFILES_BIOMATERIAL (MSS_PROFILES_BIOMATERIAL_ID, MSS_PROFILES_ID, LAB_BIOTYPE_ID, OPTION_SET)
                                    values ( @id, @profile_id, @biomat_id, case when {3} <> '' then {3} else null end )
                                    exec up_get_id 'MSS_LOAD_DICT_LOGS',1,@id out
                                    insert MSS_LOAD_DICT_LOGS (MSS_LOAD_DICT_LOGS_ID, MSS_PROFILES_ID, LAB_BIOTYPE_ID, OPTION_SET, Change)
                                    values ( @id, @profile_id, @biomat_id, {3}, 'upload' )
                                end
                             end
                        
                    ", prb.BiomatCode, prb.ProfileCode, OrgId, prb.OptionSet, prb.ProfileExtId);
            }
           
            db.Dispose();
        }
        public static void InsertTests(List<Test> Tests, int AnalyzerID)
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var test in Tests)
            {
                int success = db.ExecuteCommand(@"
                    declare @id int
                    if not exists (select code from LAB_METHODS where CODE = {0} and LAB_DEVS_ID = {2})
                    begin
                        exec up_get_id 'LAB_METHODS',1,@id out
                        insert LAB_METHODS(LAB_METHODS_ID,  CODE, LABEL, LAB_DEVS_ID)
                        values (@id, {0}, cast({1} as varchar(150)), {2})
                     end
                    ", test.Code, test.Label, AnalyzerID);

            }
            db.Dispose();
        }
        public static void InsertAuxilary(List<Invitro.Auxilary> auxilaries)
        {
            Console.WriteLine("Вставка доп. информации");
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var aux in auxilaries)
            {
                int success = db.ExecuteCommand(@"
                    declare @id int
					if not exists (select EXTERNAL_ID from MSS_AUXILARY where EXTERNAL_ID = {0})
                    begin
                        exec up_get_id 'MSS_AUXILARY',1,@id out
                        insert MSS_AUXILARY(MSS_AUXILARY_ID, EXTERNAL_ID, Name, IsRequired, MIN_VAL, MAX_VAL, Unit, ValueType, Variants)
                        values (@id, {0}, {1}, {2}, case when {3} <> '' then convert(numeric, {3}) else null end,
                             case when {4} <> '' then convert(numeric, {4}) else null end, case when {5} <> '' then {5} else null end,
                             case when {6} <> '' then {6} else null end, case when {7} <> '' then {7} else null end
                     )
                     end
                    ", aux.ExternalId, aux.Name, aux.IsRequired, aux.Min, aux.Max, aux.Unit, aux.ValType, aux.Variants);
            }
            db.Dispose();
        }

        public static void InsertDocType(List<Invitro.ExtendedDict> dicData)
        {
            Console.WriteLine("Вставка типов документов");
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var d in dicData)
            {
                int success = db.ExecuteCommand(@"
                     if not exists (select MSS_DOCTYPE_ID from MSS_DOCTYPE where MSS_DOCTYPE_ID = {0})
                     insert MSS_DOCTYPE (MSS_DOCTYPE_ID, Name)
                     values({0}, {1})
                    ", d.Id, d.Name);
            }
            db.Dispose();
        }

        public static void InsertAdresType(List<Invitro.ExtendedDict> dicData)
        {
            Console.WriteLine("Вставка типа адреса");
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var d in dicData)
            {
                int success = db.ExecuteCommand(@"
                     if not exists (select MSS_ADRESTYPE_ID from MSS_ADRESTYPE where MSS_ADRESTYPE_ID = {0})
                     insert MSS_ADRESTYPE (MSS_ADRESTYPE_ID, Name)
                     values({0},{1})
                    ", d.Id, d.Name);
            }
            db.Dispose();
        }

        public static void InsertCountry(List<Invitro.ExtendedDict> dicData)
        {
            Console.WriteLine("Вставка страны");
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var d in dicData)
            {
                int success = db.ExecuteCommand(@"
                     if not exists (select MSS_COUNTRY_ID from MSS_COUNTRY where MSS_COUNTRY_ID = {0})
                     insert MSS_COUNTRY (MSS_COUNTRY_ID, Name)
                     values({0}, {1})
                    ", d.Id, d.Name);
            }
            db.Dispose();
        }

        public static void InsertRegion(List<Invitro.ExtendedDict> dicData)
        {
            Console.WriteLine("Вставка региона");
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var d in dicData)
            {
                int success = db.ExecuteCommand(@"
                     if not exists (select MSS_REGION_ID from MSS_REGION where MSS_REGION_ID ={0})
                     insert MSS_REGION (MSS_REGION_ID, Name, ExtendedId, MSS_COUNTRY_ID, ADR_OBLAST_ID)
                     values({0}, {1}, {2}, {3}, 0)
                    ", d.Id, d.Name, d.ExtendedId, d.RefId);
            }
            db.Dispose();
        }

        public static void InsertCity(List<Invitro.ExtendedDict> dicData)
        {
            Console.WriteLine("Вставка города");
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var d in dicData)
            {
                int success = db.ExecuteCommand(@"
                     if not exists (select MSS_CITY_ID from MSS_CITY where MSS_CITY_ID = {0})
                     insert MSS_CITY (MSS_CITY_ID, Name, ExtendedId, MSS_REGION_ID, ADR_GOROD_ID)
                     values({0}, {1},  {2}, {3}, 0 )
                    ", d.Id, d.Name, d.ExtendedId, d.RefId);
            }
            db.Dispose();
        }

        public static void InsertAuxilaryProfile(List<Invitro.AuxProfile> auxProfiles, int OrgId)
        {
            Console.WriteLine("Вставка связи доп. информации с профилем");
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var ap in auxProfiles)
            {
                                int success = db.ExecuteCommand(@"
 						declare @id int, @profile_id int, @aux_id int, @biom_id int
                        if not exists (select pa.MSS_PROFILE_AUXILARY_ID 
                            from MSS_PROFILE_AUXILARY pa
                            join MSS_PROFILES p on p.MSS_PROFILES_ID = pa.MSS_PROFILES_ID
                            join MSS_AUXILARY a on a.MSS_AUXILARY_ID = pa.MSS_AUXILARY_ID
							join LAB_BIOTYPE b on b.LAB_BIOTYPE_ID = pa.LAB_BIOTYPE_ID
                            where a.EXTERNAL_ID = {0} and p.CODE = {1} and p.FM_ORG_ID = {2} and b.CODE = {3}
                                       )
                            begin
                                set @profile_id = (select top 1 MSS_PROFILES_ID from MSS_PROFILES where CODE = {1} and FM_ORG_ID = {2})
                                set @aux_id = (select top 1 MSS_AUXILARY_ID from MSS_AUXILARY where EXTERNAL_ID = {0})
								set @biom_id = (select top 1 LAB_BIOTYPE_ID from LAB_BIOTYPE where CODE = {3})
                                if @profile_id is not null and @aux_id is not null and @biom_id is not null
                                begin
	                                exec up_get_id 'MSS_PROFILE_AUXILARY',1,@id out
                                    insert MSS_PROFILE_AUXILARY (MSS_PROFILE_AUXILARY_ID, MSS_PROFILES_ID, MSS_AUXILARY_ID,  LAB_BIOTYPE_ID)
                                    values ( @id, @profile_id, @aux_id, @biom_id )
                                end
                             end
                    ", ap.AuxExternalId, ap.ProfileCode, OrgId, ap.BiomatCode);
            }
            db.Dispose();
        }

        public static void UpdateFormType()
        {
            Console.WriteLine("Test 3");
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int ok = db.ExecuteCommand(@"
                    update MSS_PROFILES_BIOMATERIAL set FORM_TYPE = ep.FORM_TYPE
                    from MSS_PROFILES_BIOMATERIAL pb
                    join LAB_BIOTYPE b on b.LAB_BIOTYPE_ID = pb.LAB_BIOTYPE_ID
                    join MSS_PROFILES p on pb.MSS_PROFILES_ID = p.MSS_PROFILES_ID
                    join MSS_EXTENDED_PROFILES ep on b.CODE = ep.BIOMATERIAL_ID and p.EXTERNAL_ID = ep.EXTERNAL_ID
                    where pb.FORM_TYPE is null
                    ");
            db.Dispose();
        }

        /*public static void UpdateFormTypelog()
        {
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            int ok = db.ExecuteCommand(@"
                    update MSS_LOAD_DICT_LOGS set FORM_TYPE = ep.FORM_TYPE
                    from MSS_LOAD_DICT_LOGS pb
                    join LAB_BIOTYPE b on b.LAB_BIOTYPE_ID = pb.LAB_BIOTYPE_ID
                    join MSS_PROFILES p on pb.MSS_PROFILES_ID = p.MSS_PROFILES_ID
                    join MSS_EXTENDED_PROFILES ep on b.CODE = ep.BIOMATERIAL_ID and p.EXTERNAL_ID = ep.EXTERNAL_ID
                    where pb.FORM_TYPE is null
                    ");
            db.Dispose();
        }*/

        public static void UpdateRequestCode(DataContext db, int RequestId, string RequestCode)
        {
            int ok = db.ExecuteCommand(@"
                        update MSS_ORDER_REQUEST set RequestCode = {1}
                        where MSS_ORDER_REQUEST_ID = {0}
                    ", RequestId, RequestCode);
        }

        public static void UpdateBarcode(DataContext db, int RequestId, string BarCode)
        {
            int ok = db.ExecuteCommand(@"
                        update MSS_BARCODE_NUMBERS set BARCODE_NUMBER = {1}
                        from MSS_BARCODE_NUMBERS brn
                        join MSS_BARCODE_SERV brs on brn.MSS_BARCODE_NUMBERS_ID = brs.MSS_BARCODE_NUMBERS_ID
                        join DIR_SERV drs on brs.DIR_SERV_ID = drs.DIR_SERV_ID
                        join MSS_ORDER_REQUEST ord on drs.MSS_ORDER_REQUEST_ID = ord.MSS_ORDER_REQUEST_ID
                        where ord.MSS_ORDER_REQUEST_ID = {0}
                    ", RequestId, BarCode);
        }

        public static void FindBiomatProfile(List<ProfileBiomat> ProfileBiomats, int OrgId)
        {
            Console.WriteLine("Проверка соответствия связи биоматериалов с профилем");
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var prb in ProfileBiomats)
            {
                
                int success = db.ExecuteCommand(@"
                         declare @id int, @profile_id int, @biomat_id int
                        if exists (select pb.MSS_PROFILES_BIOMATERIAL_ID 
                            from MSS_PROFILES_BIOMATERIAL pb
                            join MSS_PROFILES p on pb.MSS_PROFILES_ID = p.MSS_PROFILES_ID
                            join LAB_BIOTYPE b on pb.LAB_BIOTYPE_ID = b.LAB_BIOTYPE_ID
                            where b.CODE = {0} and p.EXTERNAL_ID = {4} and p.FM_ORG_ID = {2} and b.FM_ORG_ID = {2} and pb.OPTION_SET = {3}
                                       )
                            
                                /*begin
                                set @profile_id = (select top 1 MSS_PROFILES_ID from MSS_PROFILES where CODE = {1} and FM_ORG_ID = {2})
                                set @biomat_id = (select top 1 LAB_BIOTYPE_ID from LAB_BIOTYPE where CODE = {0} and FM_ORG_ID = {2})
                                if @profile_id is not null and @biomat_id is not null*/
                                begin
	                                exec up_get_id 'MSS_PROFILES_BIOMATERIAL',1,@id out
                                    update MSS_PROFILES_BIOMATERIAL 
                                    set Flag = 1
                                    where MSS_PROFILES_BIOMATERIAL_ID = (select pb.MSS_PROFILES_BIOMATERIAL_ID 
                            from MSS_PROFILES_BIOMATERIAL pb
                            join MSS_PROFILES p on pb.MSS_PROFILES_ID = p.MSS_PROFILES_ID
                            join LAB_BIOTYPE b on pb.LAB_BIOTYPE_ID = b.LAB_BIOTYPE_ID
                            where Flag = 0 and b.CODE = {0} and p.EXTERNAL_ID = {4}  and p.FM_ORG_ID = {2} and b.FM_ORG_ID = {2} and pb.OPTION_SET = {3})
                                    
                                end
                                
            
         
            ", prb.BiomatCode, prb.ProfileCode, OrgId, prb.OptionSet, prb.ProfileExtId);
}
            db.Dispose();
        }

        public static void DeleteBiomatProfilelog(List<ProfileBiomat> ProfileBiomats, int OrgId)
        {
            Console.WriteLine("Удаление связи биоматериалов с профилем ЛОГ");
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var prb in ProfileBiomats)
            {
                /*Console.WriteLine("Тест 1");*/
                int success = db.ExecuteCommand(@"
                        declare @i int, @id int, @profile_id int, @biomat_id int, @OPTION_SET varchar(50), @Change varchar(50)
                        set @id = 1
                        while (@id<50000)
                        begin
                        set @OPTION_SET =  ( select pb.OPTION_SET                              								
                        from MSS_PROFILES_BIOMATERIAL pb
                        join MSS_PROFILES p on pb.MSS_PROFILES_ID = p.MSS_PROFILES_ID
                        join LAB_BIOTYPE b on pb.LAB_BIOTYPE_ID = b.LAB_BIOTYPE_ID
                        where Flag = 0 and p.FM_ORG_ID = 975 and b.FM_ORG_ID = 975 and MSS_PROFILES_BIOMATERIAL_ID=@id)

                        set @profile_id =  ( select pb.MSS_PROFILES_ID                              								
                        from MSS_PROFILES_BIOMATERIAL pb
                        join MSS_PROFILES p on pb.MSS_PROFILES_ID = p.MSS_PROFILES_ID
                        join LAB_BIOTYPE b on pb.LAB_BIOTYPE_ID = b.LAB_BIOTYPE_ID
                        where Flag = 0 and p.FM_ORG_ID = 975 and b.FM_ORG_ID = 975 and MSS_PROFILES_BIOMATERIAL_ID=@id)

                        set @biomat_id =  ( select pb.LAB_BIOTYPE_ID                              								
                        from MSS_PROFILES_BIOMATERIAL pb
                        join MSS_PROFILES p on pb.MSS_PROFILES_ID = p.MSS_PROFILES_ID
                        join LAB_BIOTYPE b on pb.LAB_BIOTYPE_ID = b.LAB_BIOTYPE_ID
                        where Flag = 0 and p.FM_ORG_ID = 975 and b.FM_ORG_ID = 975 and MSS_PROFILES_BIOMATERIAL_ID=@id)

                        set @Change =  ( select Change                              								
                        from MSS_LOAD_DICT_LOGS
                        where MSS_PROFILES_ID = @profile_id and LAB_BIOTYPE_ID = @biomat_id and OPTION_SET = @OPTION_SET and Change='delete')

                        if @profile_id is not null and @biomat_id is not null and @OPTION_SET is not null and @Change is null
                        begin
	                        insert MSS_LOAD_DICT_LOGS (MSS_LOAD_DICT_LOGS_ID, MSS_PROFILES_ID, LAB_BIOTYPE_ID, OPTION_SET, Change)
                                                            values ( @id, @profile_id, @biomat_id, @OPTION_SET, 'delete' )
                        end

                        set @id = @id + 1
                        end
                             
                        
                    ", prb.BiomatCode, prb.ProfileCode, OrgId, prb.OptionSet, prb.ProfileExtId);

            }

            db.Dispose();
        }


        public static void DeleteBiomatProfile(List<ProfileBiomat> ProfileBiomats, int OrgId)
        {
            Console.WriteLine("Удаление связи биоматериалов с профилем");
            DataContext db = new DataContext(Processing.sqlConnectionStr);
            foreach (var prb in ProfileBiomats)
            {
                /*Console.WriteLine("Тест 1");*/
                int success = db.ExecuteCommand(@"
                        declare @id int, @profile_id int, @biomat_id int
                        if exists (select pb.MSS_PROFILES_BIOMATERIAL_ID 
                            from MSS_PROFILES_BIOMATERIAL pb
                            join MSS_PROFILES p on pb.MSS_PROFILES_ID = p.MSS_PROFILES_ID
                            join LAB_BIOTYPE b on pb.LAB_BIOTYPE_ID = b.LAB_BIOTYPE_ID
                            where Flag = 0 and p.FM_ORG_ID = 975 and b.FM_ORG_ID = 975)
                            
                            
                            begin
	                            /*exec up_get_id 'MSS_PROFILES_BIOMATERIAL',1,@id out*/
                                delete from MSS_PROFILES_BIOMATERIAL 
                                where MSS_PROFILES_BIOMATERIAL_ID = (select top 1 pb.MSS_PROFILES_BIOMATERIAL_ID 
                                    from MSS_PROFILES_BIOMATERIAL pb
                                    join MSS_PROFILES p on pb.MSS_PROFILES_ID = p.MSS_PROFILES_ID
                                    join LAB_BIOTYPE b on pb.LAB_BIOTYPE_ID = b.LAB_BIOTYPE_ID
                                    where Flag = 0 and p.FM_ORG_ID = 975 and b.FM_ORG_ID = 975)
                                
                                
                            end
                             
                        
                    ", prb.BiomatCode, prb.ProfileCode, OrgId, prb.OptionSet, prb.ProfileExtId);

            }

            db.Dispose();
        }

    }
}
