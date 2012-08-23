using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Tspd.Tspddoc;
using Word = Microsoft.Office.Interop.Word;

namespace TspdCfg.Purdue.DynTmplts
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

        public void setStyle(string styleName, Tspd.Tspddoc.TspdDocument _thisDoc, Word.Range _selRng)
        {
            try
            {
                _thisDoc.getStyleHelper().setNamedStyle(styleName, _selRng);
            }
            catch (Exception ex)
            {
               // MessageBox.Show(ex.ToString());
            }

            //Setting all formatting.
            if (_selMsg.Format.isBold)
            {
                for (int i = 1; i <= _selRng.Words.Count; i++)
                {
                    _selRng.Words[i].Bold = iTRUE;
                }
            }
            else
            {
                for (int i = 1; i <= _selRng.Words.Count; i++)
                {
                    _selRng.Words[i].Bold = iFALSE;
                }
            }

            if (_selMsg.Format.isItalics)
            {
                for (int i = 1; i <= _selRng.Words.Count; i++)
                {
                    _selRng.Words[i].Italic = iTRUE;
                }
            }
            else
            {
                for (int i = 1; i <= _selRng.Words.Count; i++)
                {
                    _selRng.Words[i].Italic = iFALSE;
                }
            }

            if (_selMsg.Format.isUnderline)
            {
                for (int i = 1; i <= _selRng.Words.Count; i++)
                {
                    _selRng.Words[i].Underline = Word.WdUnderline.wdUnderlineSingle;
                }
            }
            else
            {
                for (int i = 1; i <= _selRng.Words.Count; i++)
                {
                    _selRng.Words[i].Underline = Word.WdUnderline.wdUnderlineNone;
                }
            }

            if (_selMsg.Format.FontSize > 0)
            {
                _selRng.Font.Size = _selMsg.Format.FontSize;
            }
        }


        public void RestartNumbering(Word.Range selRng,bool flag)
        {
         
            try
            {
               // Word.ListGallery(
                object o7 = 1;
                object fal = flag;
                object applyto = Word.WdListApplyTo.wdListApplyToWholeList;
                object behaviour = Word.WdDefaultListBehavior.wdWord10ListBehavior;

                //Word.ListTemplate curTemplate = selRng.Application.ListGalleries[Word.WdListGalleryType.wdNumberGallery].ListTemplates.get_Item(ref o7);
                //curTemplate.ListLevels[1].StartAt = 1;
               // selRng.ListFormat.ListTemplate.ListLevels[1].StartAt = 1;

                selRng.ListFormat.ApplyListTemplate(selRng.ListFormat.ListTemplate, ref fal, ref applyto, ref behaviour);         

            }
            catch (Exception ex)
            {
                //
                System.Windows.Forms.MessageBox.Show("PP " + ex.ToString() + " -  " + ex.Message);
            }
           
        }

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

                    string FTbulletStyle = "";
                    string FTnumberStyle = "";
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

                                        case "FTbulletStyle":
                                            FTbulletStyle = fNode.InnerText;
                                            break;

                                        case "FTnumberStyle":
                                            FTnumberStyle = fNode.InnerText;
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
                        format = new message.formatting(paragraphStyle, bold, italics, underline, fontSize, FTbulletStyle, FTnumberStyle);

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
                private string ftBulletStyle;
                private string ftNumberStyle;

                public formatting(string _style, bool _bold, bool _italics, bool _underline,
                    int _fontSize, string _ftBulletStyle, string _ftNumberStyle)
                {
                    strStyle = _style;
                    blnBold = _bold;
                    blnItalics = _italics;
                    blnUnderline = _underline;
                    intFontSize = _fontSize;
                    ftBulletStyle = _ftBulletStyle;
                    ftNumberStyle = _ftNumberStyle;
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

                public string FtBulletStyle
                {
                    get
                    {
                        return ftBulletStyle;
                    }//end get
                }//end propery
                public string FTNumberStyle
                {
                    get
                    {
                        return ftNumberStyle;
                    }//end get
                }//end propery

            }//end class
        }//end class
    }//end class
}//end namespace