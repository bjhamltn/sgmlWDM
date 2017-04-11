using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;
using System.Net;
using System.Reflection;
using System.Collections;
using Sgml;
using System.Web;
using System.Data.OleDb;
using System.Data;
using System.Collections.Specialized;
using System.Text.RegularExpressions;
using HtmlAgilityPack;

namespace sgmlWDM
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            //string fleet = "A310_600"; MD10_WDM
            //foreach (string fleet in (new[] { "A310_300", "A310_200", "A300_600", "777", "757", "767" }))
            WdmDecoder d = new WdmDecoder();
            foreach (string fleet in (new[] { "777" }))
            {
                d.perpareSGML(@"C:\Users\795627\Desktop\" + fleet + "_WDM.sgm");
                d.setDbName("wirelist_" + fleet.Replace("_", " "));
                String wdmFile_dtd_closed = @"C:\Users\795627\Desktop\wdmFile_dtd_closed.xml";

                string[] cnames = d.decode_equipmentList(wdmFile_dtd_closed);
                d.buildDatabase_table("equipment", cnames, "equipment.xml");

                cnames = d.decode_wireList(wdmFile_dtd_closed);
                d.buildDatabase_wireList("wirelist", cnames, "wires.xml");
            }
        }


        public class WdmDecoder
        {
            public Hashtable equipmentLookUp = new Hashtable();
            public DataTable equipmentTable = new DataTable();

            private string connectionString;

            public string ConnectionString
            {
                get
                {
                    return connectionString;
                }
                set
                {
                    connectionString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source='" + value + "'";
                }
            }

            public List<string> EmptyTags = new List<string>();

            public WdmDecoder()
            {
                EmptyTags = new List<string>();
                HtmlAgilityPack.HtmlNode.ElementsFlags.Add("spanspec", HtmlElementFlag.Empty);
                HtmlAgilityPack.HtmlNode.ElementsFlags.Add("colspec", HtmlElementFlag.Empty);
                HtmlAgilityPack.HtmlNode.ElementsFlags.Add("grsymbol", HtmlElementFlag.Empty);
                HtmlAgilityPack.HtmlNode.ElementsFlags.Add("sbeff", HtmlElementFlag.Empty);
                HtmlAgilityPack.HtmlNode.ElementsFlags.Add("coceff", HtmlElementFlag.Empty);
                HtmlAgilityPack.HtmlNode.ElementsFlags.Add("deleted", HtmlElementFlag.Empty);
                HtmlAgilityPack.HtmlNode.ElementsFlags.Add("isempty", HtmlElementFlag.Empty);
                HtmlAgilityPack.HtmlNode.ElementsFlags.Add("revst", HtmlElementFlag.Empty);
                HtmlAgilityPack.HtmlNode.ElementsFlags.Add("revend", HtmlElementFlag.Empty);
                HtmlAgilityPack.HtmlNode.ElementsFlags.Add("cocst", HtmlElementFlag.Empty);
                HtmlAgilityPack.HtmlNode.ElementsFlags.Add("cocend", HtmlElementFlag.Empty);
                HtmlAgilityPack.HtmlNode.ElementsFlags.Add("hotlink", HtmlElementFlag.Empty);
                HtmlAgilityPack.HtmlNode.ElementsFlags.Add("aclist", HtmlElementFlag.Empty);
                EmptyTags = HtmlAgilityPack.HtmlNode.ElementsFlags.Where(fd => ((HtmlElementFlag)fd.Value) == HtmlElementFlag.Empty).ToList().Select(fd => fd.Key).ToList();

            }


            public void buildDatabase_table(string tableName, string[] cnmaes, string seedfile)
            {
                OleDbConnection con = new OleDbConnection(this.ConnectionString);
                OleDbCommand sqlCmd = con.CreateCommand();
                sqlCmd.Parameters.Clear();
                con.Open();
                XmlReader wires = XmlReader.Create(seedfile);



                while (wires.EOF == false)
                {
                    sqlCmd.Parameters.Clear();
                    wires.ReadToFollowing("eqrow");
                    if (wires.EOF == true)
                    {
                        break;
                    }
                    XmlReader wire = wires.ReadSubtree();
                    wires.Read();

                    List<string> clist = new List<string>();
                    List<string> vals = new List<string>();


                    string colname = "";
                    while (wire.EOF == false)
                    {

                        if (wire.NodeType == XmlNodeType.Element)
                        {
                            if (cnmaes.Contains(wire.Name) == false)
                            {
                                colname = "";
                                wire.Read();
                                continue;
                            }
                            colname = wire.Name;
                        }
                        else if (wire.NodeType == XmlNodeType.Text)
                        {
                            if (colname != "")
                            {
                                string val = wire.Value;
                                if (wire.Value == "")
                                {
                                    val = "0";
                                }

                                val = val.Trim(new char[] { ':', '-' });
                                val = val.Trim();
                                val = val == "0" ? "00" : val;
                                vals.Add(val);
                                clist.Add(colname);
                                sqlCmd.Parameters.AddWithValue("@" + colname, val);
                            }
                        }

                        wire.Read();
                    }
                    string a = String.Join(",", clist.ToArray());
                    string b = String.Join(",", clist.Select(fd => "@" + fd).ToArray());
                    string myInssertQuery = String.Format("INSERT INTO equipment ({1}) VALUES ({2})", tableName, a, b);
                    sqlCmd.CommandText = myInssertQuery;

                    try
                    {
                        sqlCmd.ExecuteNonQuery();
                    }
                    catch (OleDbException e)
                    {
                        Console.WriteLine(e.Message);
                    }

                }
                con.Close();

            }

            public void buildDatabase_wireList(string tableName, string[] cnmaes, string seedfile)
            {
                OleDbConnection con = new OleDbConnection(this.ConnectionString);
                OleDbCommand sqlCmd = con.CreateCommand();
                sqlCmd.Parameters.Clear();
                con.Open();
                XmlReader wires = XmlReader.Create("wires.xml");



                while (wires.EOF == false)
                {
                    sqlCmd.Parameters.Clear();
                    wires.ReadToFollowing("wire");
                    if (wires.EOF == true)
                    {
                        break;
                    }
                    XmlReader wire = wires.ReadSubtree();
                    wires.Read();

                    List<string> clist = new List<string>();
                    List<string> vals = new List<string>();


                    string colname = "";
                    while (wire.EOF == false)
                    {

                        if (wire.NodeType == XmlNodeType.Element)
                        {
                            if (cnmaes.Contains(wire.Name) == false)
                            {
                                colname = "";
                                wire.Read();
                                continue;
                            }
                            colname = wire.Name;
                        }
                        else if (wire.NodeType == XmlNodeType.Text)
                        {
                            if (colname != "")
                            {
                                string val = wire.Value;
                                if (wire.Value == "")
                                {
                                    val = "0";
                                }

                                vals.Add(val);
                                clist.Add(colname);
                                sqlCmd.Parameters.AddWithValue("@" + colname, val);
                            }
                        }

                        wire.Read();
                    }
                    string a = String.Join(",", clist.ToArray());
                    string b = String.Join(",", clist.Select(fd => "@" + fd).ToArray());
                    string myInssertQuery = String.Format("INSERT INTO wirelist ({0}) VALUES ({1})", a, b);
                    sqlCmd.CommandText = myInssertQuery;


                    sqlCmd.ExecuteNonQuery();
                }
                wires.Close();
                con.Close();
            }

       

            public string[] decode_equipmentList(String wdmFile_dtd_closed)
            {
                List<NameValueCollection> wireList = new List<NameValueCollection>();
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.DtdProcessing = DtdProcessing.Parse;
                settings.XmlResolver = new XmlUrlResolver();



                bool d = false;
                XmlReader xr = XmlReader.Create((new StreamReader(wdmFile_dtd_closed)).BaseStream, settings);
                d = xr.Read();
                int equipCnt = 0;
                d = xr.ReadToFollowing("wm");
                d = xr.ReadToDescendant("eqiplist");
                xr = xr.ReadSubtree();
                d = xr.ReadToDescendant("eqrow");
                StreamWriter sw = new StreamWriter("equipment.xml");
                sw.WriteLine("<equipment>");
                NameValueCollection wireData = new NameValueCollection();
                while (d == true)
                {
                    wireData = extractEquipmentInfor(xr, sw);
                    d = xr.ReadToFollowing("eqrow");
                    equipCnt++;
                }
                sw.WriteLine("</equipment>");
                sw.Close();
                xr.Close();
                return wireData.Keys.Cast<string>().ToArray();
            }

            public NameValueCollection extractEquipmentInfor(XmlReader wireRoot, StreamWriter sw)
            {
                string revdate = wireRoot.GetAttribute("revdate");
                string key = wireRoot.GetAttribute("key");


                XmlReader wireInfo = wireRoot.ReadSubtree();
                wireInfo.Read();
                wireInfo.ReadToDescendant("effect");

                string effrg = wireInfo.GetAttribute("effrg");
                string efftype = wireInfo.GetAttribute("efftype");

                NameValueCollection wireData = new NameValueCollection();
                wireInfo.Read();
                string currentNode = "";
                while (wireInfo.EOF == false)
                {
                    if (wireInfo.NodeType == XmlNodeType.Element)
                    {
                        #region MyRegion
                        currentNode = wireInfo.Name;

                        if (currentNode == "location")
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if (currentNode == "hdiagnbr")
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if (currentNode == "ein")
                        {
                            if (wireInfo.GetAttribute("type") != null)
                            {
                                wireData.Add("eq_type", wireInfo.GetAttribute("type").Trim());
                            }
                        }
                        else if (EmptyTags.Contains(currentNode))
                        {
                            wireInfo.Read();
                            continue;
                        }

                        #endregion
                    }
                    else if (wireInfo.NodeType == XmlNodeType.Text)
                    {
                        string value = wireInfo.Value.ToString().Trim();
                        value = Regex.Replace(value, @"\s{2,100}", " ");
                        if (currentNode == "position")
                        {
                            currentNode = "pos";
                            wireData.Add(currentNode, value);
                        }
                        else
                        {
                            wireData.Add(currentNode, value);
                        }

                    }
                    else if (wireInfo.NodeType == XmlNodeType.Whitespace)
                    {

                        if (currentNode == "")
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if (currentNode == "hdiagnbr")
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if (EmptyTags.Contains(currentNode))
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if (currentNode == "location")
                        {
                            wireInfo.Read();
                            continue;
                        }


                        if (wireData[currentNode] == null)
                        {
                            wireData.Add(currentNode, "");
                        }
                        else if (wireData.GetValues(currentNode).Length == 0)
                        {
                            wireData.Add(currentNode, "");
                        }
                    }

                    wireInfo.Read();
                }

                if (wireData["ein"] != null)
                {
                    string location = "";
                    if (wireData["stnnbr"] != null)
                    {
                        location = String.Format("{0}:{1}:{2}", wireData["stnnbr"].ToString(), wireData["wl"].ToString(), wireData["bl"].ToString());
                    }
                    else
                    {
                        location = wireData["zone"] == null ? "" : String.Format("{0}", wireData["zone"].ToString());
                    }

                    if (effrg != null)
                    {
                        effrg = expandeffetivity(effrg);
                    }

                    wireData.Add("location", location);
                    wireData.Add("effrg", effrg);
                    sw.WriteLine("<eqrow>");
                    foreach (string nName in wireData.Keys)
                    {
                        string val = wireData[nName] == null ? "" : wireData[nName].ToString();
                        val = HttpUtility.HtmlEncode(val.Trim());
                        sw.WriteLine(String.Format("<{0}>{1}</{0}>", nName.Trim(), val.Trim()));
                    }
                    sw.WriteLine("</eqrow>");
                }


                return wireData;
            }

            public string[] decode_wireList(String wdmFile_dtd_closed)
            {
                List<NameValueCollection> wireList = new List<NameValueCollection>();
                XmlReaderSettings settings = new XmlReaderSettings();

                settings.DtdProcessing = DtdProcessing.Parse;
                settings.XmlResolver = new XmlUrlResolver();
                XmlReader xr = XmlReader.Create((new StreamReader(wdmFile_dtd_closed)).BaseStream, settings);
                bool d = false;
                d = xr.Read();
                int wireEntryCnt = 0;
                d = xr.ReadToFollowing("wm");
                d = xr.ReadToDescendant("extwlist");
                xr = xr.ReadSubtree();
                d = xr.ReadToDescendant("extwrow");
                StreamWriter sw = null;
                StreamWriter sw_hookup = null;

            retry:
                try
                {
                    sw = new StreamWriter("wires.xml", false);
                    sw_hookup = new StreamWriter("hookup.xml", false);
                }
                catch
                {
                    System.Threading.Thread.Sleep(500);
                    goto retry;
                }

                sw_hookup.WriteLine("<hookup>");
                sw.WriteLine("<wires>");
                NameValueCollection wireData = new NameValueCollection();
                List<string> keys = new List<string>();
                while (d == true)
                {
                    wireData = extractWireInfor(xr, sw, sw_hookup);
                    d = xr.ReadToFollowing("extwrow");
                    wireEntryCnt++;
                    keys.AddRange(wireData.Keys.Cast<string>().ToArray());
                    keys = keys.Distinct().ToList();

                }

                sw.WriteLine("</wires>");
                sw_hookup.WriteLine("</hookup>");
                sw_hookup.Close();
                sw.Close();
                xr.Close();
                return keys.Distinct().ToArray();
            }

            public NameValueCollection extractWireInfor(XmlReader wireRoot, StreamWriter sw, StreamWriter sw_hookup)
            {
                string revdate = wireRoot.GetAttribute("revdate");
                string key = wireRoot.GetAttribute("key");


                XmlReader wireInfo = wireRoot.ReadSubtree();
                wireInfo.Read();
                wireInfo.ReadToDescendant("effect");

                string effrg = wireInfo.GetAttribute("effrg");
                string efftype = wireInfo.GetAttribute("efftype");

                NameValueCollection wireData = new NameValueCollection();
                wireInfo.Read();
                string currentNode = "";
                string preFix = "";
                string wireNumber = "";
                while (wireInfo.EOF == false)
                {
                    #region MyRegion
                    if (wireInfo.NodeType == XmlNodeType.Element)
                    {
                        #region MyRegion
                        currentNode = wireInfo.Name;

                        if (currentNode == "to")
                        {
                            preFix = "to_";
                            currentNode = preFix + currentNode;
                            wireInfo.Read();
                            continue;
                        }
                        else if (currentNode == "from")
                        {
                            preFix = "from_";
                            currentNode = preFix + currentNode;
                            wireInfo.Read();
                            continue;
                        }
                        else if ((EmptyTags.Contains(currentNode)))
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if (currentNode == "feedthru")
                        {
                            preFix = "feedthru_";
                            currentNode = preFix + currentNode;
                            wireInfo.Read();
                            continue;
                        }
                        else if (currentNode == "wire")
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if ((new string[] { "length", "wiretype", "fam", "wirerte", "bunddesc", "bundpnr", "modcode", "sensep" }).Contains(currentNode))
                        {
                            preFix = "";
                        }
                        currentNode = preFix + currentNode;
                        if (currentNode == "actwire")
                        {
                            wireData.Add("actwire", "true");
                        }
                        else if (currentNode == "nactwire")
                        {
                            currentNode = "actwire";
                            wireData.Add("actwire", "false");
                        }
                        else if (currentNode == "sparepin")
                        {
                            currentNode = "actwire";
                            wireData.Add("actwire", "sparepin");
                        }

                        if (currentNode.EndsWith("termnbr"))
                        {
                            string attr = wireInfo.GetAttribute("type");
                            attr = attr == null ? "" : attr.Trim();
                            wireData.Add(preFix + "termnbrtype", attr);
                        }
                        else if (currentNode.EndsWith("ein"))
                        {
                            string attr = wireInfo.GetAttribute("type");
                            attr = attr == null ? "" : attr.Trim();
                            wireData.Add(preFix + "termtype", attr);
                        }
                        else if (currentNode.EndsWith("wiretype"))
                        {
                            string attr = wireInfo.GetAttribute("wtcode");
                            attr = attr == null ? "" : attr.Trim();
                            if (wireData.GetValues("wtcode") != null)
                            {
                                wireData["wtcode"] += ("/" + attr);
                            }
                            else
                            {
                                wireData.Add("wtcode", attr);
                            }

                        }
                        #endregion
                    }
                    else if (wireInfo.NodeType == XmlNodeType.Text)
                    {

                        if (currentNode == "from_from")
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if (currentNode == "")
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if ("feedthru" == currentNode)
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if (EmptyTags.Contains(currentNode))
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if (currentNode == "to_to")
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if (currentNode == "hdiagnbr")
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if (currentNode == "wire")
                        {
                            wireInfo.Read();
                            continue;
                        }

                        string value = wireInfo.Value.ToString().Trim();
                        value = Regex.Replace(value, @"\s{2,100}", " ");
                        if (currentNode == "length")
                        {
                            float wirelength = 0;
                            float.TryParse(wireInfo.Value.ToString(), out wirelength);
                            value = wirelength.ToString().Trim();
                            if (wireData[currentNode] != null)
                            {
                                wirelength = float.Parse(wireData[currentNode]) + wirelength / 12f;
                                wireData[currentNode] = wirelength.ToString();
                            }
                            else
                            {
                                wireData.Add(currentNode, value);
                            }
                        }
                        else
                        {
                            wireData.Add(currentNode, value);
                        }

                    }
                    else if (wireInfo.NodeType == XmlNodeType.Whitespace)
                    {

                        if (currentNode == "from_from")
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if (currentNode == "")
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if ("feedthru" == currentNode)
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if (EmptyTags.Contains(currentNode))
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if (currentNode == "to_to")
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if (currentNode == "hdiagnbr")
                        {
                            wireInfo.Read();
                            continue;
                        }
                        else if (currentNode == "wire")
                        {
                            wireInfo.Read();
                            continue;
                        }


                        if (wireData[currentNode] == null)
                        {
                            wireData.Add(currentNode, "");
                        }
                        else if (wireData.GetValues(currentNode).Length == 0)
                        {
                            wireData.Add(currentNode, "");
                        }
                    }

                    wireInfo.Read();
                    #endregion
                }
                string effrgFull = "";
                if (effrg != null)
                {
                    effrgFull = expandeffetivity(effrg);
                }
                if (wireData["wirenbr"] != null)
                {
                    if (!wireData["wirenbr"].Trim().StartsWith(wireData["bundnbr"].Trim()))
                    {
                        wireNumber = wireData["bundnbr"] + "-" + wireData["wirenbr"];
                    }
                    else
                    {
                        wireNumber = wireData["wirenbr"];
                    }
                    wireNumber = wireNumber.Trim(new char[] { '-', ' ' }).Trim();
                    wireData.Add("wireno", wireNumber);
                    wireData.Add("effrg", effrgFull);
                    sw.WriteLine("<wire>");
                    foreach (string nName in wireData.Keys)
                    {
                        string val = wireData[nName] == null ? "ALL" : wireData[nName].ToString();
                        val = HttpUtility.HtmlEncode(val.Trim());
                        sw.WriteLine(String.Format("<{0}>{1}</{0}>", nName.Trim(), val.Trim()));
                    }
                    sw.WriteLine("</wire>");
                }
                else if (wireData["actwire"] == "sparepin")
                {

                    wireData.Add("effrg", effrgFull);
                    sw_hookup.WriteLine("<item>");
                    foreach (string nName in wireData.Keys)
                    {
                        string val = wireData[nName] == null ? "ALL" : wireData[nName].ToString();
                        val = HttpUtility.HtmlEncode(val.Trim());
                        sw_hookup.WriteLine(String.Format("<{0}>{1}</{0}>", nName.Trim(), val.Trim()));
                    }
                    sw_hookup.WriteLine(String.Format("<{0}>{1}</{0}>", "status", "unused"));
                    sw_hookup.WriteLine("</item>");
                }


                return wireData;
            }

            public void setDbName(string dbName)
            {
                if (File.Exists(dbName + ".mdb"))
                {
                    File.Delete(dbName + ".mdb");
                }
                File.Copy("wirelist_template.mdb", dbName + ".mdb");
                string dsd = Path.GetFullPath(dbName + ".mdb");
                this.ConnectionString = Path.GetFullPath(dbName + ".mdb");
            }

            public void perpareSGML(string wdmFile)
            {

                String wdmFile_dtd = @"C:\Users\795627\Desktop\wdmFile_dtd.xml";
                String wdmFile_dtd_closed = @"C:\Users\795627\Desktop\wdmFile_dtd_closed.xml";

                StreamReader sr = new StreamReader(wdmFile);
                StreamWriter sw = new StreamWriter(wdmFile_dtd, false);
                sw.WriteLine(dtd.entity_def4);
                char[] buff = new char[1000];
                while (sr.EndOfStream == false)
                {
                    int len = sr.Read(buff, 0, buff.Length);
                    sw.Write(buff, 0, len);
                }
                sw.Close();
                sr.Close();
                closeSelfClosing(wdmFile_dtd, wdmFile_dtd_closed);


            }


            void closeSelfClosing(string filepath_in, string filepath_out)
            {

                StreamReader strR = new StreamReader(filepath_in);
                using (StreamWriter strW = new StreamWriter(filepath_out, false))
                {
                    while (!strR.EndOfStream)
                    {
                        Console.Write("\r" + (100 * strR.BaseStream.Position / strR.BaseStream.Length).ToString());
                        GC.Collect();
                        List<bool> boos = Enumerable.Repeat(true, 2550).ToList();
                        List<string> lines = boos.Select(fd => !strR.EndOfStream ? strR.ReadLine() : "").ToList();
                        string block = string.Join("\r\n", lines).Trim();
                        block = block.Replace("\r\n>", " >");
                        foreach (string tagname in EmptyTags)
                        {
                            block = Regex.Replace(block, @"<(?<a>" + tagname + "[^>]*?)>", delegate(Match m)
                            {
                                if (m.Groups["a"].Value.Trim().EndsWith("/"))
                                {
                                    return m.Value;
                                }
                                else
                                {
                                    string sd = string.Format("<{0}/>", m.Groups["a"].Value);
                                    return sd;
                                }
                            }, RegexOptions.IgnoreCase);
                        }
                        strW.WriteLine(block);
                    }
                    strR.Close();
                    strW.Close();
                }
            }


            string expandeffetivity(string effrg)
            {
                List<string> range = effrg.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries).ToList();

                MatchCollection rangeUnits = Regex.Matches(effrg, @"(?<a>\d{3})(?<b>\d{3})");

                List<int> tails = rangeUnits.Cast<Match>().SelectMany(fd => Enumerable.Range(int.Parse(fd.Groups["a"].Value), 1 + int.Parse(fd.Groups["b"].Value) - int.Parse(fd.Groups["a"].Value))).ToList();
                return String.Join(",", tails);
            }

        }

    }
}
