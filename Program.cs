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

namespace sgmlWDM
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            //string fleet = "A310_600"; MD10_WDM
            //foreach (string fleet in (new[] { "A310_300", "A310_200", "A300_600", "777", "757", "767" }))
            foreach (string fleet in (new[] { "MD11", "MD10" }))
            {
                WdmDecoder d = new WdmDecoder();
                d.perpareSGML(@"C:\Users\795627\Desktop\" + fleet + "_WDM.sgm");
                d.setDbName("wirelist_" + fleet.Replace("_", " "));
                String wdmFile_dtd_closed = @"C:\Users\795627\Desktop\wdmFile_dtd_closed.xml";

                string[] cnames = d.decode_equipmentList(wdmFile_dtd_closed);
                d.buildDatabase_table("equipment", cnames, "equipment.xml");

                cnames = d.decode_wireList(wdmFile_dtd_closed);
                d.buildDatabase_wireList("wirelist", cnames, "wires2.xml");
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
                    connectionString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source='" + value+"'";
                }
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
                XmlReader wires = XmlReader.Create("wires2.xml");

                //string[] cnmaes ={"actwire",
                //                "wirecode",
                //                "wirenbr",
                //                "wireawg",
                //                "wireclr",
                //                "bundnbr",
                //                "length",
                //                "wtcode",
                //                "wiretype",
                //                "refint",
                //                "fam",
                //                "bundpnr",
                //                "bunddesc",
                //                "sensep",
                //                "from_termtype",
                //                "from_ein",
                //                "from_termnbrtype",
                //                "from_termnbr",
                //                "from_shunt",
                //                "from_termcode",
                //                "to_termtype",
                //                "to_ein",
                //                "to_termnbrtype",
                //                "to_termnbr",
                //                "to_shunt",
                //                "to_termcode",
                //                "modcode","wireno",
                //                "effrg"};

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
                            if(cnmaes.Contains(wire.Name)== false)
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

            protected void AutoCloseElementsInternal(SgmlReader reader, XmlWriter writer)
            {

                object msgBody = reader.NameTable.Add("MSGBODY");

                object previousElement = null;
                Stack elementsWeAlreadyEnded = new Stack();
                int idx = 0;
                Stack openElements = new Stack();

                {
                    while (reader.Read())
                    {
                        idx++;
                        switch (reader.NodeType)
                        {
                            case XmlNodeType.Element:
                                previousElement = reader.LocalName;
                                writer.WriteStartElement(reader.LocalName.ToLower());
                                if (reader.HasAttributes)
                                {
                                    reader.MoveToFirstAttribute();
                                }
                                for (int attr_idx = 0; attr_idx < reader.AttributeCount; attr_idx++)
                                {
                                    
                                    string attrName = reader.Name.ToLower();
                                    string attrValue = reader.Value.Trim();
                                    writer.WriteAttributeString(attrName, attrValue);                                    
                                    if(reader.MoveToNextAttribute() == false)
                                    {
                                        break;
                                    }
                                }
                                    
                                openElements.Push(previousElement);
                                break;
                            case XmlNodeType.Text:
                                if (openElements.Count > 0)
                                {
                                    if (String.IsNullOrEmpty(reader.Value)==false)
                                    {

                                        string dd = reader.Value.Trim();
                                        byte s = (byte)dd[0];
                                        if (s < 32)
                                        {
                                            Console.WriteLine("Skip Char " + dd);
                                            writer.WriteString("...");
                                        }
                                        else
                                        {
                                            HttpUtility.HtmlEncode(dd);
                                            writer.WriteString(dd);

                                        }
                                        if (previousElement != null && !previousElement.Equals(msgBody))
                                        {

                                            writer.WriteEndElement();
                                            elementsWeAlreadyEnded.Push(previousElement);
                                            openElements.Pop();
                                        }
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("big problems?");
                                }
                                break;
                            case XmlNodeType.EndElement:
                                if (elementsWeAlreadyEnded.Count > 0
                                    && Object.ReferenceEquals(elementsWeAlreadyEnded.Peek(),
                                       reader.LocalName))
                                {
                                    elementsWeAlreadyEnded.Pop();
                                }
                                else
                                {
                                    writer.WriteEndElement();
                                    openElements.Pop();
                                }
                                break;
                            case XmlNodeType.Whitespace:
                                break;
                            default:
                                writer.WriteNode(reader, false);
                                break;
                        }
                    }
                }

            }
            
            public string[] decode_equipmentList(String wdmFile_dtd_closed)
            {
                List<NameValueCollection> wireList = new List<NameValueCollection>();
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.DtdProcessing = DtdProcessing.Parse;
                settings.XmlResolver = new XmlUrlResolver();
                XmlReader xr = XmlReader.Create((new StreamReader(wdmFile_dtd_closed)).BaseStream, settings);
                bool d = false;
                d = xr.Read();
                int equipCnt = 0;
                d = xr.ReadToDescendant("eqiplist");
                xr = xr.ReadSubtree();
                d = xr.ReadToDescendant("eqrow");
                StreamWriter sw = new StreamWriter("equipment.xml");
                sw.WriteLine("<equipment>");
                NameValueCollection wireData = new NameValueCollection();
                while (d == true)
                {
                   wireData  = extractEquipmentInfor(xr, sw);
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
                d = xr.ReadToDescendant("extwlist");
                xr = xr.ReadSubtree();
                d = xr.ReadToDescendant("extwrow");
                //StreamWriter sw = new StreamWriter("wires.xml",false);
                StreamWriter sw = null;
                
            retry:
                try
                {
                    sw = new StreamWriter("wires2.xml", false);
                }
                catch
                {
                    System.Threading.Thread.Sleep(500);
                    goto retry;
                }
                
                sw.WriteLine("<wires>");
                NameValueCollection wireData = new NameValueCollection();
                List<string> keys = new List<string>();
                while (d == true)
                {                    
                    wireData = extractWireInfor(xr, sw);
                    d = xr.ReadToFollowing("extwrow");
                    wireEntryCnt++;
                    keys.AddRange(wireData.Keys.Cast<string>().ToArray());
                    keys = keys.Distinct().ToList() ;

                }
                sw.WriteLine("</wires>");
                sw.Close();
                xr.Close();
                return keys.Distinct().ToArray();
            }

            public NameValueCollection extractWireInfor(XmlReader wireRoot, StreamWriter sw)
            {
                string revdate = wireRoot.GetAttribute("revdate");
                string key = wireRoot.GetAttribute("key");


                XmlReader wireInfo = wireRoot.ReadSubtree();
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
                        else if ((new string[] { "revst", "revend", "" }.Contains(currentNode)))
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
                        else if ((new string[] { "length", "wiretype", "fam", "wirerte?", "bunddesc", "bundpnr", "modcode", "sensep" }).Contains(currentNode))
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
                            if (wireData.GetValues("wtcode")!= null)
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
                        else if ((new string[] { "revst", "revend", "", "feedthru" }.Contains(currentNode)))
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
                        else if ((new string[] { "revst", "revend", "", "feedthru" }.Contains(currentNode)))
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
                    wireData.Add("effrg", effrg);
                    sw.WriteLine("<wire>");
                    foreach (string nName in wireData.Keys)
                    {
                        string val = wireData[nName] == null? "ALL" : wireData[nName].ToString();
                        val = HttpUtility.HtmlEncode(val.Trim());
                        sw.WriteLine(String.Format("<{0}>{1}</{0}>", nName.Trim(), val.Trim()));
                    }
                    sw.WriteLine("</wire>");
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
                    int len = sr.Read(buff,0,buff.Length );
                    sw.Write(buff, 0, len);
                }
                sw.Close();


                XmlReaderSettings settings = new XmlReaderSettings();
                settings.DtdProcessing = DtdProcessing.Parse;
                settings.XmlResolver = new XmlUrlResolver();
                Sgml.SgmlReader sgmlReader = new Sgml.SgmlReader();
                sgmlReader.InputStream = new StreamReader(wdmFile_dtd);

                XmlReader xr = SgmlReader.Create((new StreamReader(wdmFile_dtd)).BaseStream, settings);

                XmlWriterSettings xwset = new XmlWriterSettings();
                xwset.Indent = true;
                xwset.IndentChars = "\t";
                xwset.NewLineChars = "\r\n";
                xwset.ConformanceLevel = ConformanceLevel.Auto;
                XmlWriter xw = XmlWriter.Create(wdmFile_dtd_closed, xwset);
                AutoCloseElementsInternal(sgmlReader, xw);
                xw.Close();

               


            }
        }
    }
}
