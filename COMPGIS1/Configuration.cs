using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml; 
using System.IO;

namespace COMPGIS1
{
    public class Configuration
    {
        
        private string _configfilepath = "";
        
        private XmlDataDocument xmldoc = new XmlDataDocument();
        private XmlNodeList xmlnode;

        // Line Section Display Items
        private int _bplsdisplaymode = 0;
        private string _bpsavedcolorrampname = "";
        private int _bpriskfieldtouse = 0;

        public string FilePath
        {
            get { return _configfilepath; }
            set { _configfilepath = value; }
        }

        public int BPLSDisplayMode
        {
            get { return _bplsdisplaymode; }
            set { _bplsdisplaymode = value; }
        }

        public string BPSavedColorRampName
        {
            get { return _bpsavedcolorrampname; }
            set { _bpsavedcolorrampname = value; }
        }

        public int BPRiskFieldtoUse
        {
            get { return _bpriskfieldtouse; }
            set { _bpriskfieldtouse = value; }
        }


        //-------------------------------------
        //-------------------------------------
        //-------------------------------------
        public void LoadConfigFile()
        {
            if (_configfilepath.Length == 0) return;
            if (!File.Exists(_configfilepath))
                SaveConfigFile();
            else
            {
                try
                {
                    // Read the File In
                    XmlReader xmlReader = XmlReader.Create(_configfilepath);
                    while (xmlReader.Read())
                    {
                        // BP Display Mode
                        if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "BP"))
                        {
                            if (xmlReader.HasAttributes)
                                _bplsdisplaymode = Convert.ToInt32(xmlReader.GetAttribute("displaymode"));
                        }
                        // BP Saved Color Ramp
                        if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "BP"))
                        {
                            if (xmlReader.HasAttributes)
                                _bpsavedcolorrampname = xmlReader.GetAttribute("savedcolorramp");
                        }
                    }
                    xmlReader.Close();
                }
                catch (Exception)
                {
                }
            }
        }

        //-------------------------------------
        //-------------------------------------
        //-------------------------------------
        public void SaveConfigFile()
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                XmlNode rootNode = xmlDoc.CreateElement("Configuration");
                xmlDoc.AppendChild(rootNode);

                // Buried Piping
                XmlNode xmlNode = xmlDoc.CreateElement("BP");

                // BP Display Mode
                XmlAttribute attribute = xmlDoc.CreateAttribute("displaymode");
                attribute.Value = _bplsdisplaymode.ToString();
                xmlNode.Attributes.Append(attribute);

                // BP Saved Color Ramp
                attribute = xmlDoc.CreateAttribute("savedcolorramp");
                attribute.Value = _bpsavedcolorrampname;
                xmlNode.Attributes.Append(attribute);

                // BP Risk Field To Use
                attribute = xmlDoc.CreateAttribute("riskfield");
                attribute.Value = _bpriskfieldtouse.ToString();
                xmlNode.Attributes.Append(attribute);

                // Done - Save Data
                rootNode.AppendChild(xmlNode);
                xmlDoc.Save(_configfilepath);                
            }
            catch (Exception)
            {                
            }
        }

    }

}


