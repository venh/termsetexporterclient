using System;
using System.Net;
using System.Linq;
using System.Xml.Linq;
using System.Collections;
using System.Globalization;
using System.Configuration;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using TermSetExporterClient.TaxonomyClientSvcProxy;

namespace TermSetExporterClient
{
    class Program
    {
        #region Static Variables

        static string SITE_URL = string.Empty;
        static string COLUMN_GROUP_NAME = string.Empty;
        static string TERMS_XML_FILE_NAME = string.Empty;
        static string HIERARCHY_NEEDED = string.Empty;
        static string USER_NAME = string.Empty;
        static string PASSWORD = string.Empty;
        static string DOMAIN = string.Empty;

        #endregion

        static void Main(string[] args)
        {
            XDocument fldsDoc = null;
            IEnumerable<XElement> fldEles = null;
            XDocument termSetDoc = null;
            ConsoleKeyInfo key;
            try
            {
                #region Variable Initialization

                SITE_URL = ConfigurationManager.AppSettings["SITE_URL"];
                COLUMN_GROUP_NAME = ConfigurationManager.AppSettings["COLUMN_GROUP_NAME"];
                TERMS_XML_FILE_NAME = ConfigurationManager.AppSettings["TERMS_XML_FILE_NAME"];
                HIERARCHY_NEEDED = ConfigurationManager.AppSettings["HIERARCHY_NEEDED"];               
                Console.Write("Please enter a valid user name: ");
                USER_NAME = Console.ReadLine();
                Console.Write("Please enter the password: ");               
                do
                {
                    key = Console.ReadKey(true);
                    // Backspace Should Not Work
                    if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                    {
                        PASSWORD = string.Concat(PASSWORD, key.KeyChar);
                        Console.Write("*");
                    }
                    else
                    {
                        if (key.Key == ConsoleKey.Backspace && PASSWORD.Length > 0)
                        {
                            PASSWORD = PASSWORD.Substring(0, (PASSWORD.Length - 1));
                            Console.Write("\b \b");
                        }
                    }
                }
                // Stops Receving Keys Once Enter is Pressed
                while (key.Key != ConsoleKey.Enter);
                Console.WriteLine();
                Console.Write("Please enter a valid domain: ");
                DOMAIN = Console.ReadLine();

                #endregion

                Console.WriteLine("Term Set Xml Generation Started...");
                using (ClientContext ctxt = new ClientContext(SITE_URL))
                {
                    ctxt.Credentials = new NetworkCredential(USER_NAME, PASSWORD, DOMAIN);
                    Web web = ctxt.Web;
                    FieldCollection fields = web.Fields;
                    ctxt.Load(fields, flds => flds.SchemaXml);
                    ctxt.ExecuteQuery();
                    fldsDoc = XDocument.Parse(fields.SchemaXml);
                    fldEles = from fldEle in fldsDoc.Root.Descendants().OfType<XElement>()
                              where ((fldEle.Attribute("Type") != null) && (fldEle.Attribute("Group") != null) && (fldEle.Attribute("Group").Value == COLUMN_GROUP_NAME) &&
                              ((fldEle.Attribute("Type").Value == "TaxonomyFieldTypeMulti") || (fldEle.Attribute("Type").Value == "TaxonomyFieldType")))
                              select fldEle;
                    termSetDoc = GetOutputTermSetsNode();
                    if ((fldEles != null) && (fldEles.Count() > 0))
                    {
                        foreach (XElement fieldEle in fldEles)
                        {
                            string tsXml = GetTermSetXML(fieldEle.ToString());
                            if (bool.Parse(HIERARCHY_NEEDED.ToUpperInvariant()))
                            {
                                termSetDoc.Root.Add(GetOutputTermSetDocument(tsXml));
                            }
                            else
                            {
                                AddTermSet(termSetDoc, tsXml);
                            }
                        }
                    }
                }
                termSetDoc.Save(TERMS_XML_FILE_NAME);
                Console.WriteLine("Term Set Xml Generation Completed Successfully...");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadLine();
            }
            finally
            {
                fldEles = null;
                fldsDoc = null;
                termSetDoc = null;              
            }
        }

        #region TermSet Retriever Functions

        static string GetTermSetXML(string schemaXml)
        {
            Taxonomywebservice ws = null;
            string tsId = string.Empty;
            string sspId = string.Empty;
            string termSetXML = string.Empty;
            try
            {
                ws = new Taxonomywebservice(); // Implement Singleton here. <TBD>
                if (SITE_URL.EndsWith("/"))
                {
                    ws.Url = string.Concat(SITE_URL, "_vti_bin/TaxonomyClientService.asmx");
                }
                else
                {
                    ws.Url = string.Concat(SITE_URL, "/_vti_bin/TaxonomyClientService.asmx");
                }
                //ws.UseDefaultCredentials = true;
                ws.Credentials = new NetworkCredential(USER_NAME, PASSWORD, DOMAIN);
                string version = "<version>0</version>";
                string oldTimeStamp = "<timeStamp>-1</timeStamp>";
                string timeStamp = string.Empty;
                sspId = GetTermSetOrSSPID(schemaXml, "SspId");
                tsId = GetTermSetOrSSPID(schemaXml, "TermSetId");
                termSetXML = ws.GetTermSets("<termStoreId>" + sspId + "</termStoreId>", "<termSetId>" + tsId + "</termSetId>", CultureInfo.CurrentUICulture.LCID, oldTimeStamp, version, out timeStamp);
                return termSetXML;
            }
            catch
            {
                throw;
            }
            finally
            {
                if (ws != null)
                {
                    ws.Abort();
                    ws.Dispose();
                    ws = null;
                }
            }
        }

        static string GetTermSetOrSSPID(string schemaXML, string entity)
        {
            string id = string.Empty;
            XElement schemaEle = null;
            IEnumerable<XElement> propEles = null;
            try
            {
                schemaEle = XElement.Parse(schemaXML);
                propEles = from pEle in schemaEle.Descendants("Property")
                           where pEle.Descendants("Name").ElementAt(0).Value == entity
                           select pEle;
                if (propEles.Count() == 1)
                {
                    id = propEles.ElementAt(0).Descendants("Value").ElementAt(0).Value;
                }
                return id;
            }
            catch
            {
                throw;
            }
            finally
            {
                schemaEle = null;
                propEles = null;
            }
        }

        #endregion

        #region Hierarchy Functions

        static XElement GetOutputTermSetElement(XDocument tSetDoc)
        {
            IEnumerable<XElement> tsEles = null;
            try
            {
                tsEles = from tsEle in tSetDoc.Root.Descendants().OfType<XElement>()
                         where tsEle.Name == "TS"
                         select tsEle;
                string tsName = tsEles.ElementAt(0).Attribute("a12").Value;
                string tsID = tsEles.ElementAt(0).Attribute("a9").Value;
                return new XElement("TermSet", new XAttribute("Name", tsName), new XAttribute("ID", tsID));
            }
            catch
            {
                throw;
            }
            finally
            {
                tsEles = null;
            }
        }

        static XElement GetParentTermElements(XDocument tSetDoc, XElement termSetEle, string tsId)
        {
            IEnumerable<XElement> ptEles = null;
            try
            {
                ptEles = from tEle in tSetDoc.Root.Descendants().OfType<XElement>()
                         where ((tEle.Name == "T") && (tEle.Descendants("TM").Count() > 0)
                         && (tEle.Descendants("TM").ElementAt(0).Attribute("a25") == null))
                         select tEle;
                foreach (XElement ptEle in ptEles)
                {
                    string termId = ptEle.Attribute("a9").Value;
                    string termName = ptEle.Descendants("TL").ElementAt(0).Attribute("a32").Value;                  
                    termSetEle.Add(new XElement("Term", new XAttribute("Name",termName),new XAttribute("ID",termId)));
                }
                return termSetEle;
            }
            catch
            {
                throw;
            }
            finally
            {
                ptEles = null;
            }
        }

        static XElement GetChildTermElements(XDocument tSetDoc, XElement ptEle)
        {
            IEnumerable<XElement> childEles = null;
            try
            {
                childEles = from cEle in tSetDoc.Root.Descendants().OfType<XElement>()
                            where ((cEle.Name == "T") && (cEle.Descendants("TM").Count() > 0)
                            && (cEle.Descendants("TM").ElementAt(0).Attribute("a25") != null))
                            select cEle;
                foreach (XElement childEle in childEles)
                {
                    string pTermID = childEle.Descendants("TM").ElementAt(0).Attribute("a25").Value;
                    string termId = childEle.Attribute("a9").Value;
                    if (GetParent(ptEle, pTermID) != null)
                    {
                        string termName = childEle.Descendants("TL").ElementAt(0).Attribute("a32").Value;
                        ptEle.Descendants("Term").Where(item => item.Attribute("ID").Value == pTermID).FirstOrDefault()
                         .AddFirst(new XElement("Term", new XAttribute("Name", termName), new XAttribute("ID", termId)));
                    }
                }
                return ptEle;
            }
            catch
            {
                throw;
            }
            finally
            {
                childEles = null;
            }
        }

        static XDocument GetOutputTermSetsNode()
        {
            XDocument termDoc = null;
            XElement rootEle = null;
            try
            {
                termDoc = new XDocument();
                rootEle = new XElement("TermSets");
                termDoc.AddFirst(rootEle);
                return termDoc;
            }
            catch
            {
                throw;
            }
            finally
            {
                termDoc = null;
                rootEle = null;
            }
        }

        static XElement GetOutputTermSetDocument(string ipTSXml)
        {
            XElement opTSEle = null;
            XDocument ipTSDoc = null;
            try
            {
                ipTSDoc = XDocument.Parse(ipTSXml);
                opTSEle = GetOutputTermSetElement(ipTSDoc);
                opTSEle = GetParentTermElements(ipTSDoc, opTSEle, opTSEle.Attribute("ID").Value);
                opTSEle = GetChildTermElements(ipTSDoc, opTSEle);
                return opTSEle;
            }
            catch
            {
                throw;
            }
            finally
            {
                opTSEle = null;
                ipTSDoc = null;
            }
        }

        static XElement GetParent(XElement opTermsEle, string id)
        {
            IEnumerable<XElement> ptEles = null;
            try
            {
                ptEles = from ptEle in opTermsEle.Descendants()
                         where ptEle.Attribute("ID").Value == id
                         select ptEle;
                if (ptEles.Count() == 1)
                {
                    return ptEles.ElementAt(0);
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                ptEles = null;
            }
        }

        #endregion

        #region Non Hierarchy Functions

        static XDocument CreateTermSetDocument()
        {
            XDocument termDoc = null;
            XElement rootEle = null;
            try
            {
                termDoc = new XDocument();
                rootEle = new XElement("TermSets");
                termDoc.AddFirst(rootEle);
                return termDoc;
            }
            catch
            {
                throw;
            }
            finally
            {
                termDoc = null;
                rootEle = null;
            }
        }

        static void AddTermSet(XDocument termSetsDoc, string termSetXml)
        {
            XDocument tsDoc = null;
            IEnumerable<XElement> tsEles = null;
            IEnumerable<XElement> tEles = null;
            XElement tsElement = null;
            try
            {
                tsDoc = XDocument.Parse(termSetXml);
                tsEles = from tsEle in tsDoc.Root.Descendants().OfType<XElement>()
                         where tsEle.Name == "TS"
                         select tsEle;
                if (tsEles.Count() == 1)
                {
                    string tsName = tsEles.ElementAt(0).Attribute("a12").Value;
                    string tsID = tsEles.ElementAt(0).Attribute("a9").Value;
                    tsElement = new XElement("TermSet", new XAttribute("Name", tsName), new XAttribute("ID", tsID));
                    tEles = from tEle in tsEles.ElementAt(0).Parent.Descendants().OfType<XElement>()
                            where tEle.Name == "T"
                            select tEle;
                    foreach (XElement term in tEles)
                    {
                        string termId = term.Attribute("a9").Value;
                        string termName = term.Descendants("TL").ElementAt(0).Attribute("a32").Value;
                        tsElement.Add(new XElement("Term", new XAttribute("Name", termName), new XAttribute("ID", termId)));
                    }
                    termSetsDoc.Root.Add(tsElement);
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                tsDoc = null;
                tsEles = null;
                tsElement = null;
            }
        }

        #endregion
    }
}
