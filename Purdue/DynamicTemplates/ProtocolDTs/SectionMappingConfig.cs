using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

using Tspd.Utilities;

namespace TspdCfg.Purdue.DynTmplts
{
    public class SectionMappingConfig
    {
        private XmlNode xTaskStyles;
        public message _selMsg = null;
        private Dictionary<string, List<message>> myStyleList = new Dictionary<string, List<message>>();



        public SectionMappingConfig(XmlNode _taskStyleNode)
        {
            xTaskStyles = _taskStyleNode;
            loadStyles();
        }//end function

        public message getMessageByName(string _styleName, string _lineName)
        {
            foreach (KeyValuePair<string, List<message>> kp in myStyleList)
            {
                if (kp.Key == _styleName)
                {
                    foreach(message m in kp.Value)
                        if (m.Name == _lineName)
                        {
                            _selMsg = m;
                            return m;
                        }
                }//end if
            }//end foreach

            return null;
        }//end function

        private void loadStyles()
        {
            foreach (XmlNode styleNode in xTaskStyles)
            {
                List<message> lstMessages = new List<message>();
                foreach (XmlNode x in styleNode.ChildNodes)
                {
                    switch (x.Name)
                    {
                        case "Message":
                            loadItem(x, lstMessages);
                            break;
                    }//end switch

                }//end foreach

                myStyleList.Add(styleNode.Attributes.GetNamedItem("name").Value, lstMessages);
            }//end foreach
        }//end function

        private void loadItem(XmlNode _Node, List<message> lstMessages)
        {
            XmlNode msg = _Node;

            string name = msg.Attributes.GetNamedItem("name").InnerText;

            string paragraphStyle = "Normal";
            bool bold = false;
            bool italics = false;
            bool underline = false;
            bool newline = false;  
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
                                case "NewLine":
                                    success = bool.TryParse(fNode.InnerText, out newline);
                                    if (!success) newline = false;
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
                format = new message.formatting(paragraphStyle, bold, italics, underline, fontSize,newline);
            lstMessages.Add(new message(name, text, format));

        }//end function

        public void setStyle(string styleName, Tspd.Tspddoc.TspdDocument _thisDoc, Word.Range _selRng)
        {
            try
            {
                if (styleName !=null && styleName.Length >0)
                {
                    _thisDoc.getStyleHelper().setNamedStyle(styleName, _selRng);
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.ToString());
            }



            //Setting all formatting.
            if (_selMsg.Format.isBold)
            {
                _selRng.Bold = VBAHelper.iTRUE;
            }
            else
            {
                _selRng.Bold = VBAHelper.iFALSE;
            }

            if (_selMsg.Format.isItalics)
            {
                _selRng.Italic = VBAHelper.iTRUE;
            }
            else
            {
                _selRng.Italic = VBAHelper.iFALSE;
            }

            if (_selMsg.Format.isUnderline)
            {
                _selRng.Underline = Word.WdUnderline.wdUnderlineSingle;
            }
            else
            {
                _selRng.Underline = Word.WdUnderline.wdUnderlineNone;
            }
        }


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
                private bool blnNewLine;

                public formatting(string _style, bool _bold, bool _italics, bool _underline,
                    int _fontSize,bool _newline)
                {
                    strStyle = _style;
                    blnBold = _bold;
                    blnItalics = _italics;
                    blnUnderline = _underline;
                    intFontSize = _fontSize;
                    blnNewLine = _newline;
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

                public bool isNewLine
                {
                    get
                    {
                        return blnNewLine;
                    }//end get
                }//end property
            }//end class
        }//end class
    }//end class
}//end namespace
