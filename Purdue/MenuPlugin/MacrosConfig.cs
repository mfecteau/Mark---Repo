using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Tspd.Tspddoc;


namespace TspdCfg.FastTrack.PlugIn
{
    public class MacrosConfig
    {
        private string strConfigXMLPath;
        private string strMacroElementPath;
        private List<message> lstMessages = new List<message>();
        public message _selMsg = null;
        static public readonly int iTRUE = -1;
        static public readonly int iFALSE = 0;

        public MacrosConfig(string _configXMLPath, string _macroElementPath)
        {
            strConfigXMLPath = _configXMLPath;
            strMacroElementPath = _macroElementPath;
            setProperties();
        }//end function

        public message getMessageByName(string _name)
        {
            foreach (message m in lstMessages)
            {
                if (m.Name == _name)
                {
                    _selMsg = m;
                    return m;
                }
            }//end foreach

            return null;
        }//end function

        

        private void setProperties()
        {
            XmlNode macroNode = getNode();

            foreach (XmlNode x in macroNode.ChildNodes)
            {
                switch (x.Name)
                {
                    case "Messages":
                        loadList(x);
                        break;
                }//end switch
                
            }//end foreach
        }//end function

        private void loadList(XmlNode _Node)
        {
            foreach (XmlNode msg in _Node.ChildNodes)
            {
                if (msg.Name.Equals("Message"))
                {
                    string name = msg.Attributes.GetNamedItem("name").InnerText;

                    string paragraphStyle = "Normal";
                    bool bold = false;
                    bool italics = false;
                    bool underline = false;
                    int fontSize = 12;
                    string text = "";

                    bool success = false;
                    bool hasFormatting = false;

                    foreach (XmlNode cNode in msg.ChildNodes)
                    {
                        switch (cNode.Name)
                        {
                            case "Formatting":
                                hasFormatting = true;
                                foreach (XmlNode fNode in cNode.ChildNodes)
                                {
                                    switch (fNode.Name)
                                    {
                                        case "Bold":
                                            success = bool.TryParse(fNode.InnerText, out bold);
                                            if (!success) bold = false;
                                            break;
                                        case "Italics":
                                            success = bool.TryParse(fNode.InnerText, out italics);
                                            if (!success) italics = false;
                                            break;
                                        case "Underline":
                                            success = bool.TryParse(fNode.InnerText, out underline);
                                            if (!success) underline = false;
                                            break;
                                        case "FontSize":
                                            try { fontSize = int.Parse(fNode.InnerText); }//end try
                                            catch { fontSize = 12; }//end catch
                                            break;
                                        case "ParagraphStyle":
                                            paragraphStyle = fNode.InnerText;
                                            break;
                                    }//end switch
                                }//end foreach
                                break;
                            case "Text":
                                text = cNode.InnerText;
                                break;
                        }//end switch
                    }//end foreach

                    message.formatting format = null;

                    if (hasFormatting)
                        format = new message.formatting(paragraphStyle, bold, italics, underline, fontSize);

                    lstMessages.Add(new message(name, text, format));
                }//end If
            }//end foreach
        }//end function

        private XmlNode getNode()
        {
            XmlDocument myDoc = new XmlDocument();
            myDoc.Load(strConfigXMLPath);

            XmlNodeList myMacroList = myDoc.GetElementsByTagName("Macro");

            foreach (XmlNode macroNode in myMacroList)
            {
                if (macroNode.Attributes.GetNamedItem("elementPath").InnerText == strMacroElementPath)
                {
                    return macroNode;
                }//end if
            }//end foreach

            throw new Exception("Macro configuration not found");
        }//end function

        public List<message> Messages
        {
            get
            {
                return lstMessages;
            }//end get
        }//end function

        public class message
        {
            private string strName;
            private string strText;
            private formatting clsFormat;

            public message(string _name, string _message, formatting _format)
            {
                strName = _name;
                strText = _message;
                clsFormat = _format;
            }//end function

            public string Name
            {
                get
                {
                    return strName;
                }//end get
            }//end property

            public string Text
            {
                get
                {
                    return strText;
                }//end get
            }//end property

            public formatting Format
            {
                get
                {
                    return clsFormat;
                }//end get
            }//end property

            public class formatting
            {
                private string strStyle;
                private bool blnBold;
                private bool blnItalics;
                private bool blnUnderline;
                private int intFontSize;

                public formatting(string _style, bool _bold, bool _italics, bool _underline,
                    int _fontSize)
                {
                    strStyle = _style;
                    blnBold = _bold;
                    blnItalics = _italics;
                    blnUnderline = _underline;
                    intFontSize = _fontSize;
                }//end function

                public string Style
                {
                    get
                    {
                        return strStyle;
                    }//end get
                }//end propery

                public bool isBold
                {
                    get
                    {
                        return blnBold;
                    }//end get
                }//end property

                public bool isItalics
                {
                    get
                    {
                        return blnItalics;
                    }//end get
                }//end property

                public bool isUnderline
                {
                    get
                    {
                        return blnUnderline;
                    }//end get
                }//end property

                public int FontSize
                {
                    get
                    {
                        return intFontSize;
                    }//end get
                }//end property
            }//end class
        }//end class
    }//end class
}//end namespace
