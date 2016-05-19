using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using SageWSSelDataClassLibrary.SageWSWebReferenceSyracuse;
//using SageWSSelDataClassLibrary.ServiceReferenceSageWS;
using System.Text.RegularExpressions;
using System.Net;
using System.Web.Services.Protocols;
using SageWSSelDataClassLibrary;

namespace SageWSSelData
{

    public class CallWebServiceX3
    {
        public static int MAX_NB_COLUMN = 300;
        private CAdxCallContext cAdxCallContext;

        public CAdxCallContext CAdxCallContext
        {
            get { return cAdxCallContext; }
            set { cAdxCallContext = value; }
        }
        private CAdxWebServiceXmlCCService cAdxWebServiceXmlCCService;

        public CAdxWebServiceXmlCCService CAdxWebServiceXmlCCService
        {
            get { return cAdxWebServiceXmlCCService; }
            set { cAdxWebServiceXmlCCService = value; }
        }
        private XmlDocument xmlDocument;

        public XmlDocument XmlDocument
        {
            get { return xmlDocument; }
            set { xmlDocument = value; }
        }

        private CAdxResultXml cAdxResultXml;

        public CAdxResultXml CAdxResultXml
        {
            get { return cAdxResultXml; }
            set { cAdxResultXml = value; }
        }

        private String inputXmlString;

        public String InputXmlString
        {
            get { return inputXmlString; }
            //set { inputXmlString = value; }
        }

        public CallWebServiceX3(String url, CAdxCallContext cAdxCallContext, string user, string password)
        {
            this.cAdxCallContext = cAdxCallContext;
            // cAdxWebServiceXmlCCService = (MyWebService)new MyWebService();
            //cAdxWebServiceXmlCCService = new CAdxWebServiceXmlCCService();
            cAdxWebServiceXmlCCService = new MyWebService();
            //cAdxWebServiceXmlCCService = new MyWebService();
            cAdxWebServiceXmlCCService.Url = url;
            //url = "http://52.17.90.248:8124/soap-generic/syracuse/collaboration/syracuse/CAdxWebServiceXmlCC";
            //NetworkCredential netCredential = new NetworkCredential(cAdxCallContext.codeUser, cAdxCallContext.password);
            NetworkCredential netCredential = new NetworkCredential(user, password);
            cAdxWebServiceXmlCCService.Credentials = netCredential;
            //Uri uri = new Uri(url);

            //ICredentials credentials = netCredential.GetCredential(uri, "Basic");

            //cAdxWebServiceXmlCCService.Credentials = credentials;

            //cAdxWebServiceXmlCCService.PreAuthenticate = true;

            xmlDocument = new XmlDocument();
            this.cAdxResultXml = null;
            this.inputXmlString = "";
            //this.cAdxCallContext.requestConfig = "adxwss.trace.on=on&adonix.trace.on=on";
            //this.cAdxCallContext.requestConfig = "adxwss.trace.on=on";
        }



        public void SelData(String publicName, String type, int nextt, String[] tabcrit, ref int allsel, ref int nbsel,
            ref int tabnb, ref List<String> listvalstr, ref int staret, ref String mesret, ref int noReq)
        {

            //String inputXmlString;
            String resultxml;
            XmlNodeList xmlNodeList;
            StringBuilder inputXmlSB;
            inputXmlSB = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF - 8\"?><PARAM><GRP ID=\"GRP1\" ><FLD NAME=\"TYP\">");
            inputXmlSB.Append(type);
            inputXmlSB.Append("</FLD><FLD NAME=\"NEXTT\">");
            inputXmlSB.Append(nextt);
            inputXmlSB.Append("</FLD></GRP><TAB ID=\"GRP2\">");
            inputXmlString = inputXmlSB.ToString();
            for (int i = 0; i < tabcrit.Length; i++)
            {
                //if (!tabcrit[i].Equals(""))
               // {
                inputXmlSB.Append("<LIN NUM=\"");
                inputXmlSB.Append(i + 1);
                inputXmlSB.Append("\"><FLD NAME=\"TABCRIT\">");
                inputXmlSB.Append(tabcrit[i]);
                inputXmlSB.Append("</FLD></LIN>");
                //}
            }
            inputXmlSB.Append("</TAB>");
            inputXmlSB.Append("<GRP ID=\"GRP3\" ><FLD NAME=\"NOREQ\">");
            inputXmlSB.Append(noReq);
            inputXmlSB.Append("</FLD></GRP>");
            inputXmlSB.Append("</PARAM>");
            inputXmlString = inputXmlSB.ToString();
            try
            {

                cAdxResultXml = cAdxWebServiceXmlCCService.run(cAdxCallContext, publicName, inputXmlString);
                resultxml = cAdxResultXml.resultXml;
                if (cAdxResultXml.status == 0)
                {
                    staret = 100;
                    if (cAdxResultXml.messages.Length >= 1)
                    {
                        mesret = cAdxResultXml.messages[0].message;
                    }
                }
                else
                {
                    xmlDocument.LoadXml(resultxml);
                    xmlNodeList = xmlDocument.SelectNodes("/RESULT/GRP/FLD");
                    // Parcourt l'ensemble des noeuds éléments
                    foreach (XmlNode objNode in xmlNodeList)
                    {
                        if (objNode.Attributes.GetNamedItem("NAME").Value.Equals("TABNB"))
                        {
                            tabnb = int.Parse(objNode.InnerText);
                        }
                        if (objNode.Attributes.GetNamedItem("NAME").Value.Equals("ALLSEL"))
                        {
                            allsel = int.Parse(objNode.InnerText);
                        }

                        if (objNode.Attributes.GetNamedItem("NAME").Value.Equals("NBSEL"))
                        {
                            nbsel = int.Parse(objNode.InnerText);
                        }
                        if (objNode.Attributes.GetNamedItem("NAME").Value.Equals("STARET"))
                        {
                            staret = int.Parse(objNode.InnerText);
                        }
                        if (objNode.Attributes.GetNamedItem("NAME").Value.Equals("MESRET"))
                        {
                            mesret = objNode.InnerText;
                        }
                        if (objNode.Attributes.GetNamedItem("NAME").Value.Equals("NOREQ"))
                        {
                            noReq = int.Parse(objNode.InnerText);
                        }
                    }
                    xmlNodeList = xmlDocument.SelectNodes("/RESULT/TAB/LIN/FLD");
                    // Parcourt l'ensemble des noeuds éléments
                    listvalstr.Clear();
                    foreach (XmlNode objNode in xmlNodeList)
                    {
                        if (objNode.Attributes.GetNamedItem("NAME") != null && objNode.Attributes.GetNamedItem("NAME").Value.Equals("TABVALCLB"))
                        {
                            listvalstr.Add(objNode.InnerText);
                            //Console.Out.WriteLine("- toto" + objNode.InnerXml);
                        }
                    }
                }

            }
            catch (Exception e)
            {
                staret = 200;
                mesret = e.Message;
            }
        }

        public void UpdData(String publicName, String type, string filt, ref int nbupd, string transac, int nbtabclb, int nbkey, int nbval,
            String[] tabkeyclb, String[] tabvalclb, ref int staret, ref String mesret)
        {

            //String inputXmlString;
            String resultxml;
            XmlNodeList xmlNodeList;
            StringBuilder inputXmlSB;

            inputXmlSB = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF - 8\"?><PARAM><GRP ID=\"GRP1\" ><FLD NAME=\"TYP\">");
            inputXmlSB.Append(type);
            inputXmlSB.Append("</FLD><FLD NAME=\"FILT\">");
            inputXmlSB.Append(filt);
            inputXmlSB.Append("</FLD><FLD NAME=\"NBUPD\">");
            inputXmlSB.Append(nbupd);
            inputXmlSB.Append("</FLD><FLD NAME=\"TRANSAC\">");
            inputXmlSB.Append(transac);
            inputXmlSB.Append("</FLD><FLD NAME=\"NBTABCLB\">");
            inputXmlSB.Append(nbtabclb);
            inputXmlSB.Append("</FLD><FLD NAME=\"NBKEY\">");
            inputXmlSB.Append(nbkey);
            inputXmlSB.Append("</FLD><FLD NAME=\"NBVAL\">");
            inputXmlSB.Append(nbval);
            inputXmlSB.Append("</FLD></GRP><TAB ID=\"GRP2\">");
            for (int i = 0; i < nbtabclb; i++)
            {
                if (!tabkeyclb[i].Equals(""))
                {
                    //inputXmlSB.Append("<LIN><FLD NAME=\"TABKEYCLB\">");
                    inputXmlSB.Append("<LIN NUM=\"");
                    inputXmlSB.Append(i + 1);
                    inputXmlSB.Append("\"><FLD NAME=\"TABKEYCLB\">");
                    inputXmlSB.Append(tabkeyclb[i]);
                    inputXmlSB.Append("</FLD>");
                    if (!tabvalclb[i].Equals(""))
                    {
                        inputXmlSB.Append("<FLD NAME=\"TABVALCLB\">");
                        inputXmlSB.Append(tabvalclb[i]);
                        inputXmlSB.Append("</FLD></LIN>");
                    }
                }
            }
            inputXmlSB.Append("</TAB></PARAM>");
            inputXmlString = inputXmlSB.ToString();
            inputXmlString = inputXmlString.Replace("&", "&amp;");
            try
            {
                cAdxResultXml = cAdxWebServiceXmlCCService.run(cAdxCallContext, publicName, inputXmlString);
                resultxml = cAdxResultXml.resultXml;
                if (cAdxResultXml.status == 0)
                {
                    staret = 100;
                    if (cAdxResultXml.messages.Length >= 1)
                    {
                        mesret = cAdxResultXml.messages[0].message;
                    }
                }
                else
                {
                    xmlDocument.LoadXml(resultxml);
                    xmlNodeList = xmlDocument.SelectNodes("/RESULT/GRP/FLD");
                    // Parcourt l'ensemble des noeuds éléments
                    foreach (XmlNode objNode in xmlNodeList)
                    {
                        if (objNode.Attributes.GetNamedItem("NAME").Value.Equals("NBUPD"))
                        {
                            nbupd = int.Parse(objNode.InnerText);
                        }
                        if (objNode.Attributes.GetNamedItem("NAME").Value.Equals("STARET"))
                        {
                            staret = int.Parse(objNode.InnerText);
                        }
                        if (objNode.Attributes.GetNamedItem("NAME").Value.Equals("MESRET"))
                        {
                            mesret = objNode.InnerText;
                        }

                    }
                }
            }
            catch (Exception e)
            {
                staret = 200;
                mesret = e.Message;
            }
        }

        private String[,] getTab2Param(List<String> listvalstr, int tabnb, ref int nbtabval)
        {

            String[,] ret = new String[tabnb, MAX_NB_COLUMN];
            Regex regex;
            String[] tabString;
            int col = -1, lig = -1;
            String mot = "";
            foreach (var valstr in listvalstr)
            {
                if (valstr != "")
                {
                    lig += 1;
                    col = -1;
                    regex = new Regex("{");
                    tabString = regex.Split(valstr);
                    regex = new Regex("\"");
                    tabString = regex.Split(tabString[1]);
                    // Console.Out.WriteLine(" tabString--" + tabString[0]);
                    for (int i = 0; i < tabString.Length; i++)
                    {
                        if (i != 0) // en 0 il ya un "" à ne pas récupérer
                        {
                            mot = tabString[i];
                            if (!mot.Equals(",") && !mot.Equals("}"))
                            {
                                //    Console.Out.WriteLine(" mot--" + mot);
                                col += 1;
                                nbtabval = col + 1;
                                ret[lig, col] = mot;
                            }
                        }
                    }

                }
            }
            return ret;
        }



        public void SelData(String publicName, String type, int nextt, String[] tabcrit, ref int allsel, ref int nbsel,
            ref int tabnb, ref int nbtabval, out String[,] tabvalstr, ref int staret, ref String mesret, ref int noReq)
        {
            List<String> listvalstr = new List<string>();
            SelData(publicName, type, nextt, tabcrit, ref allsel, ref nbsel, ref tabnb, ref listvalstr, ref staret, ref mesret, ref noReq);
            tabvalstr = getTab2Param(listvalstr, tabnb, ref nbtabval);
            //allsel = 2;
        }
    }
}
