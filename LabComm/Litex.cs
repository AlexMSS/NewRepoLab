using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;



namespace LabComm
{
    public class Litex
    {
        public class LabData
        {
            public string Code;
            public string Label;
            public string CodeBiom;
            public string LabelBiom;
            public string Property;
        }
        public class RegData
        {
            public string Code;
            public string Label;
            public string ExternalId;
            public string TestsLine;
        }

        public static void LoadDictionaries()
        {
            XDocument xdoc;
            string[] dicfiles = Directory.GetFiles(Processing.Ini.DictFilePath, "*.xml");
            XNamespace ns = "urn:schemas-microsoft-com:office:spreadsheet";
            XNamespace ss = "urn:schemas-microsoft-com:office:spreadsheet";
            List<LabData> labDataList = new List<LabData>();
            List<RegData> regDataList = new List<RegData>();
            foreach (string f in dicfiles)
            {
                xdoc = XDocument.Load(f);
                if (f.IndexOf("LAB") > -1)
                {
                    var row = xdoc.Element(ns + "Workbook").Element(ns + "Worksheet").Elements(ns + "Table");
                    foreach (var cell in row.Elements(ns + "Row"))
                    {
                        LabData labData = new LabData();
                        if (cell.Element(ns + "Cell").Attribute(ns + "Index") != null)
                        {
                            var data = (from x in cell.Elements(ns + "Cell")  //.Elements(ns + "Data")
                                        select x.Element(ns + "Data").Value).ToList(); //x.Attribute(ns + "Index").Value).ToList();
                            string gr = data[10].Substring(0, 2);
                            labData.Code = data[0];
                            labData.Label = data[8];
                            labData.CodeBiom = data[2] + data[4] + gr;
                            labData.LabelBiom = data[3] + "|" + data[5] + (gr == "00" ? "" : "|" + gr);
                            labDataList.Add(labData);
                        }
                    }

                }

                if (f.IndexOf("REG") > -1)
                {
                    var row = xdoc.Element(ns + "Workbook").Element(ns + "Worksheet").Elements(ns + "Table");
                    foreach (var cell in row.Elements(ns + "Row"))
                    {
                        RegData regData = new RegData();
                        if (cell.Element(ns + "Cell").Attribute(ns + "Index") != null)
                        {
                            var data = (from x in cell.Elements(ns + "Cell")  //.Elements(ns + "Data")
                                        select x.Element(ns + "Data").Value).ToList(); //x.Attribute(ns + "Index").Value).ToList();
                            regData.Code = data[0];
                            regData.Label = data[1];
                            regData.ExternalId = data[2];
                            regData.TestsLine = data[3];
                            regDataList.Add(regData);
                        }
                    }
                }

                var biomat = labDataList.GroupBy(g => new { g.CodeBiom, g.LabelBiom })
                            .Select(b => new Biomat { Code = b.Key.CodeBiom, Label = b.Key.LabelBiom }).Where(b => b.Code != "").Take(3).ToList(); //.Where(b=> b.Code == "" || b.Code == "")
                DBControl.InsertBiomaterial(biomat, 978);
                var profiles = regDataList.Select(p => new Profile { Code = p.Code, Label = p.Label, ExternalId = p.ExternalId }).Where(p => p.Code != "").Take(3).ToList();
                //DBControl.InsertProfiles(profiles, 978);
                List<ProfileBiomat> profileBiomatList = new List<ProfileBiomat>();
                foreach(var p in regDataList.Take(3))
                {
                    var testCodes = p.TestsLine.Split(new char[] { '|' });
                    string codeBiom = "";
                    ProfileBiomat profileBiomat = new ProfileBiomat();
                    foreach(string code in testCodes)
                    {
                        int index = labDataList.FindIndex(s => string.Equals(s.Code, code));
                        if (index != -1)
                            if (labDataList[index].CodeBiom != codeBiom)
                            {
                                profileBiomat.ProfileCode = p.Code;
                                codeBiom = labDataList[index].CodeBiom;
                                profileBiomat.BiomatCode = codeBiom;
                                profileBiomatList.Add(profileBiomat);
                            }
                    }
                }
                //DBControl.InsertBoamatProfile(profileBiomatList, 978);

                var tests = labDataList.Select(t => new Test { Code = t.Code, Label = t.Label }).Where(t => t.Code != "").Take(3).ToList();
                DBControl.InsertTests(tests, 978);
                // foreach (var d in profiles)  Console.WriteLine("{0}\t{1}\t{2}\t{3}", d.Code, d.Label, "", ""); //d.ExternalId, d.TestsLine
            }

        }
    }
}
