using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections;
using MSXML2;
using System.Xml;
using System.IO;
using System.Windows.Forms;

using Tspd.Businessobject;
using Tspd.Tspddoc;
using Tspd.Icp;
using Tspd.Utilities;
using Tspd.Rules;
using Tspd.Context;
using Tspd.Bridge;
using WorX;


namespace TspdCfg.FastTrack.Rules
{
    public class ModDocSection : ITSPDRule
	{
#if false
<rule id="ModDocSection" type="CSHARP" displayName="ModifiedDocsection" source="TspdCfg.FastTrack.Rules.ModDocSection.,Rules.dll" categories="testcs" debug="false"/>
#endif
        public static Hashtable Sel_Section = new Hashtable();
        public static readonly string MY_ID = "ModDocSection";

		private static string _ruleId = "";
		private static bool _debug = false;

		public void Init(string ruleId, bool debug) 
		{
			_ruleId = ruleId;
			_debug = debug;
		}

		public string AdvisoryPrefix
		{
			get { return MY_ID; }
		}

        public bool canRunInStandaloneDocument
        {
            get { return true; }

        }

		public ICollection Run() 
		{
			ArrayList advisories = new ArrayList();
			ArrayList docSec =  new ArrayList();

			ContextManager ctx = ContextManager.getInstance();
			TspdDocument doc = ctx.getActiveDocument();
			BusinessObjectMgr bom = doc.getBom();
            string filepath = doc.getTrialProject().getTemplateDirPath() + "\\rules\\RulesConfig.xml";
            Hashtable htLibItems =  ReadConfigXML(filepath);
			docSec = CompareDocSectionTEXT(doc,htLibItems);
         
			string smsg="" ;
		
			foreach (string str in docSec) 
            {
				 smsg = str;
			RuleAdvisory adv = new RuleAdvisory(
				_ruleId, MY_ID + str , smsg);
			advisories.Add(adv);
			}
			

			return advisories;
		}

		public ArrayList CompareDocSectionTEXT(TspdDocument currDoc,Hashtable htLibItems)
		{
		
			//Get the Docuement Sections cannot exceed Level 4.

			ArrayList myList2 = new ArrayList();
            string filePath = null,_selRngText="",plainText="";

            string tempPath = Path.GetTempPath();
           // string newFilePath = "";

            RichTextBox rtBox = new RichTextBox();
			IEnumerator sections = currDoc.getDocSectionList();
			
			while(sections.MoveNext())
			{
				DocumentSectionEntry dse = (DocumentSectionEntry)sections.Current;
				IElementBase worxEntry = (IElementBase)dse.getWorxElement();

                if (htLibItems.ContainsKey(dse.getElementPath().ToLower()))
                    if (dse.getDocumentState() == Tspd.Businessobject.ChooserEntry.DocumentState.InDoc)
                    {
                        LibraryElement libEle = (LibraryElement)htLibItems[dse.getElementPath().ToLower()];
                        if (libEle.getContentType() == LibraryContentType.MSWORD)
                        {
                         //GET FILE PATH 
                            try
                            {
                                filePath = BridgeProxy.getInstance().loadLibraryElement(libEle.getLibraryBucketID(), libEle.getPKValue());
                            }
                            catch (Exception e)
                            {
                                Log.exception(e, libEle.getActualDisplayValue() + " - library item not found!");
                            }


                        //Handling RTF & WORD
                            if (System.IO.Path.GetExtension(filePath) == ".rtf")
                            {
                              try
                                {
                                    rtBox.LoadFile(filePath);                                    
                                    plainText = rtBox.Text;                                   
                                }
                                catch (Exception e)
                                {
                                    Log.exception(e, e.Message);
                                }
                            }
                            else if (System.IO.Path.GetExtension(filePath) == ".doc")
                            {
                                Log.trace(dse.getActualDisplayValue() + " has a word document(.doc) as its library item. Please contact Library Adminstrator for an rtf file.");

                            }

                             _selRngText = worxEntry.WdRange.Text.Trim();
                             _selRngText = _selRngText.Replace("\r", "\n");

                             if (_selRngText.Length > dse.getActualDisplayValue().Length + 1)
                             {
                                 _selRngText = _selRngText.Substring(dse.getActualDisplayValue().Length + 1);
                             }
                             else
                             {
                                 _selRngText = "";
                             }

                             _selRngText = CleanTags(_selRngText);
                             plainText = CleanTags(plainText);

                            if (!_selRngText.Equals(plainText.Trim()))
                            {
                                myList2.Add(dse.getActualDisplayValue() + " is a standard document section, and has been changed.");
                            }
                            plainText = "";  //Clear the parameter, after each loop.
                        }
                    } //END IF InDoc
                			
			} //END WHILE
			
			return myList2;
		}


        private string CleanTags(string str)
        {
            str = str.Replace("\n", "");  //new line
            str = str.Replace("\t", "");  //tab 
            str = str.Replace("\a", "");  //new line break
            str = str.Replace("\r", "");  //carraige return
            str = str.Replace("\v", "");  //vertical tab
            str = str.Replace("\f", "");  //form feed
            return str;
        }
        public Hashtable ReadConfigXML(string path)
        {
            Hashtable htLibItems = new Hashtable();
            try
            {
                if (System.IO.File.Exists(path))
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(path);

                    string libItemname = "";
                    string libCategory = "";
                    //Get Library Item Bucket
                    LibraryManager lm = LibraryManager.getInstance();
                    LibraryElement lbItem = null;

                    // Select and display all Tasks.
                    XmlNodeList nodeList;
                    XmlElement root = doc.DocumentElement;
                    nodeList = root.SelectNodes("/Advisory ");
                 //   XmlNode on
                    foreach (XmlNode xNode in nodeList)
                    {
                        foreach (XmlNode xiNode in xNode.ChildNodes)
                        {
                            libItemname = xiNode.Attributes.GetNamedItem("libraryItem").Value;
                            libCategory = xiNode.Attributes.GetNamedItem("category").Value;

                            lbItem = lm.getLibraryElement(libCategory, libItemname);

                            if (lbItem != null)
                            {
                                htLibItems.Add(xiNode.Attributes.GetNamedItem("path").Value.ToLower(), lbItem);  //Adding node to ArrayList.
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Log.exception(e, e.Message + " Modified Doc Section: Config file is not present");
            }

            return htLibItems;
        }

	}
}
