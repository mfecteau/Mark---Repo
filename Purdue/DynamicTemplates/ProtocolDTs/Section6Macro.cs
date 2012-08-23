using System;
using System.Collections;
using System.Linq;

using Tspd.Tspddoc;
using Tspd.MacroBase;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;
using Tspd.Bridge;

using System.IO;
using System.Xml;
using System.Windows.Forms;

using Word = Microsoft.Office.Interop.Word;

namespace VersionControl 
{
    internal sealed class Section6Macro
	{
        private static readonly string header_ = @"$Header: Section6Macro.cs, 1, 10.12.2010 04:10, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
    /// <summary>
    /// Summary description for Section6 Macro.
    /// </summary>
    public class Section6Macro : AbstractMacroImpl
    {
        SOA _currentSOA = null;
        TaskListClass tl = null;
        MacrosConfig mc = null;
        SectionMappingConfig secConfig = null;
        SectionInsertMethods insertmethod = null;
       
        string _docsectionName = "";
        long _selectedObjID = 0;

        Hashtable htSectionlookup = new Hashtable();


        XmlNode ParentTaskNode = null;
        XmlNode docsectionNode = null;
        XmlNode stylenode = null;
        XmlNode taskListnode = null;
        XmlNode rootICPNode = null;

        LibraryManager lm = null;  //Library Manager.

        public Section6Macro(MacroExecutor.MacroParameters mp)
            : base(mp)
        {
        }

        #region Dynamic Tmplt Methods

        #region Section6

        public static MacroExecutor.MacroRetCd DisplaySection(
            MacroExecutor.MacroParameters mp)
        {
#if false
<ChooserEntry elementPath="TspdCfg.AZ.DynTmplts.Section6Macro.DisplaySection,ProtocolDTs.dll" elementLabel="Section 6" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Schedule" autogenerates="true" toolTip="Section5" shouldRun="true">

</ChooserEntry>
#endif
            try
            {
                mp.pba_.setOperation("Inserting Section Builder", "Generating information...");

                Section6Macro macro = null;
                macro = new Section6Macro(mp);
                macro.preProcess();
                macro.display();
                macro.postProcess();
                return macro.macroStatusCode_;
            }
            catch (Exception e)
            {
                Log.exception(e, "Error in Section Builder Macro");
                mp.inoutRng_.Text = "Variables by Section Builder  Macro: " + e.Message;
            }
            return MacroExecutor.MacroRetCd.Failed;
        }

        #endregion

        #endregion

 

        public override void preProcess()
        {
            #region Macro Config
            try
            {
                string chooserElementPath = this.macroEntry_.getElementPath();
                string fPath = tspdDoc_.getTrialProject().getTemplateDirPath() + "\\dyntmplts\\MacrosConfig.xml";
                mc = new MacrosConfig(fPath, chooserElementPath);
            }
            catch (Exception e)
            {
                Log.exception(e, e.Message + e.StackTrace);
            }
            #endregion

            #region Parameters
            // Get stored parameters
            string sParms = execParms_.getParm(MacroExecutor.MacroExecParameters.PARM_1);
            string aParms = null;

            if (!MacroBaseUtilities.isEmpty(sParms))
            {
                aParms = sParms;
            }

            bool parmsValid = true;

            if (aParms != null)
            {
                if (!MacroBaseUtilities.isEmpty(aParms))
                {
                    _docsectionName = aParms;

                }
            }
            else
            {
                parmsValid = false;
            }

            // Ask the user if the parms are missing/invalid
            if (!parmsValid)
            {
                _docsectionName = "";
            }
            #endregion
        }

        public override void display()
        {
            string msg = "";
            Word.Range inoutRange = this.startAtBeginningOfParagraph();
            Word.Range wrkRng = inoutRange.Duplicate;


            if (BuildTaskList(wrkRng))
            {
                if (PrintDocumentSection6(docsectionNode, wrkRng))
                {

                    if (inoutRange.End == wrkRng.End)
                    {
                        msg = mc.getMessageByName("exception2").Text;
                        msg = msg.Replace("[[macroname]]", ParentTaskNode.Attributes.GetNamedItem("name").Value);
                        wrkRng.InsertAfter(msg);
                        wrkRng.InsertParagraphAfter();
                        mc.setStyle(mc.getMessageByName("exception2").Format.Style, tspdDoc_, wrkRng);
                        wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

                    }

                    execParms_.setParm(MacroExecutor.MacroExecParameters.PARM_1,_docsectionName);
                    inoutRange.End = wrkRng.End;
                    setOutgoingRng(inoutRange);
                    wdDoc_.UndoClear();
                    this.MacroStatusCode = MacroExecutor.MacroRetCd.Succeeded;
                    pba_.updateProgress(1.0);
                    return;
                }
            }

            //If Build and Print both returns true, means successfull else it will come over here as failed
            inoutRange.End = wrkRng.End;
            setOutgoingRng(inoutRange);
            wdDoc_.UndoClear();
            this.MacroStatusCode = MacroExecutor.MacroRetCd.Failed;

            pba_.updateProgress(1.0);
            return;
        }


        #region ReadingXML

        public void readXML()
        {

            #region configFile
            try
            {
                string chooserElementPath = this.macroEntry_.getElementPath();
                string fPath = tspdDoc_.getTrialProject().getTemplateDirPath() + "\\dyntmplts\\SectionElementMapping.xml";
                XmlDocument myDoc = new XmlDocument();
                myDoc.Load(fPath);

                XmlNodeList myMacroList = myDoc.GetElementsByTagName("Macro");

                if (myMacroList.Count > 0)
                {
                    InsertDocSection frmDocsec = new InsertDocSection();
                    frmDocsec.ParseNodeListforSectionSelection(myMacroList);

                    if (frmDocsec.ShowDialog() == DialogResult.OK)
                    {

                    }
                    else
                    {
                    }

                }

                // ParentTaskNode = getNode(fPath, chooserElementPath);

            }
            catch (Exception e)
            {
                Log.exception(e, e.Message + e.StackTrace);
                return;
            }
            #endregion

        }

        public bool BuildTaskList(Word.Range wrkRng)
        {

            #region configFile
            try
            {
                string chooserElementPath = this.macroEntry_.getElementPath();
                string fPath = tspdDoc_.getTrialProject().getTemplateDirPath() + "\\dyntmplts\\SectionElementMapping.xml";
                ParentTaskNode = getNode(fPath, chooserElementPath, _docsectionName);

            }
            catch (Exception e)
            {
                Log.exception(e, e.Message + e.StackTrace);
                return false;
            }
            #endregion


            if (ParentTaskNode == null)
            {
                return false; // Print error message and exit out of DCO.
            }


            //Inititializing the LibraryManager:
            lm = LibraryManager.getInstance();


            tl = new TaskListClass();
            tl._bom = bom_;
            tl._icdSchemaMgr = this.icdSchemaMgr_;

            string name, tasktype, tasklab;
            ArrayList arrvalue = new ArrayList();
            ArrayList TaskswithEpochorPurpose = null;

            bool readDocSec = false;


            foreach (XmlNode macroNode in ParentTaskNode.ChildNodes)
            {
                if (macroNode.Name.ToLower().StartsWith("tasklists"))
                {
                    taskListnode = macroNode;
                }
                else if (macroNode.Name.ToLower().StartsWith("documentsection"))
                {
                    docsectionNode = macroNode;
                }
                else if (macroNode.Name.ToLower().StartsWith("taskstyles"))
                {
                    stylenode = macroNode;

                    if (stylenode != null)
                    {
                        //for Task Styles.
                        secConfig = new SectionMappingConfig(stylenode);
                    }

                }
            } //ENDFOR

            //Initializing the ICP Root Node
            rootICPNode = bom_.getIcp().getRoot();


            foreach (XmlNode xiChildNode in taskListnode.ChildNodes)
            {
                name = xiChildNode.Attributes.GetNamedItem("name").Value;
                tasktype = xiChildNode.Attributes.GetNamedItem("type").Value;
                tasklab = xiChildNode.Attributes.GetNamedItem("labs").Value;
                arrvalue = new ArrayList();
                foreach (XmlNode xivalNode in xiChildNode.ChildNodes)
                {
                    arrvalue.Add(xivalNode.InnerText);
                }

                //Empty ArrayList
                TaskswithEpochorPurpose = new ArrayList();

                //Get ArrayList of task based on it epoch or task-visit purpose
                tl.AddItem(name, tasktype, tasklab, arrvalue, TaskswithEpochorPurpose);
            }
            //    //Assuming that first node will always be for "TaskList" and next would always be "Document Section".
            //    readDocSec = true;

            //} //End For

            //Most Imp Call for DCO. this will initialize
            if (taskListnode.ChildNodes.Count > 0)
            {
                tl.FillTaskList();
            }
            //  ReadDocSectionsNode(docsectionNode);
            Get_Child_of_Parent(docsectionNode, false);

            if (ParentTaskNode.Attributes.GetNamedItem("sectionstart") != null)
            {
                getSectionNumber(docsectionNode.SelectNodes("DocumentSection"), ParentTaskNode.Attributes.GetNamedItem("sectionstart").Value);
            }

            return true;
        }

        //All the methods for Reading Xml.
        private XmlNode getNode(string strConfigXMLPath, string chooserPath, string _storedNodevalue)
        {

            XmlDocument myDoc = new XmlDocument();
            myDoc.Load(strConfigXMLPath);
            XmlNodeList myMacroList = myDoc.GetElementsByTagName("Macro");

            if (_storedNodevalue.Length <= 0)
            {
                //Call Form
                InsertDocSection ObjDocSection = new InsertDocSection();
                ObjDocSection.ParseNodeListforSectionSelection(myMacroList);

                if (ObjDocSection.DialogResult == DialogResult.OK)
                {
                    _storedNodevalue = ObjDocSection._sectionName;
                    _docsectionName = _storedNodevalue;
                }
                else
                {
                    return null;
                }
            }




            Log.trace("Selected Section:: " + _storedNodevalue);

            foreach (XmlNode macroNode in myMacroList)
            {
                if (macroNode.Attributes.GetNamedItem("name").InnerText == _storedNodevalue)
                {
                    Log.trace("Found");
                    return macroNode;
                }//end if
            }//end foreach

            throw new Exception("Macro configuration not found");
        }//end function

        #endregion


        #region DocumentSection
        public void ReadDocSectionsNode(XmlNode docSectionsnode)
        {
            try
            {
                foreach (XmlNode cnode in docSectionsnode.ChildNodes)
                {
                    // MessageBox.Show(cnode.Attributes.GetNamedItem("label").Value + cnode.Attributes.GetNamedItem("preserve").Value);
                    foreach (XmlNode pnode in cnode.ChildNodes)
                    {
                        // MessageBox.Show(pnode.Attributes.GetNamedItem("label").Value + pnode.Attributes.GetNamedItem("preserve").Value);
                        foreach (XmlNode gcnode in pnode.ChildNodes)
                        {
                            MessageBox.Show(gcnode.Attributes.GetNamedItem("type").Value);
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString() + ex.StackTrace);
            }
        }


        /*  Method "Get_Child_of_Parent" is Main Method. It is a RECCURSIVE LOOP, which will go through each child and check if it has value.
         *          * 
         */
        public bool Get_Child_of_Parent(XmlNode pNode, bool _isPrint)
        {
            try
            { 
                foreach (XmlNode cNode in pNode)
                {
                    if (cNode.Name.ToLower() == "documentsection")
                    {
                        Log.trace(cNode.Name + "::" + cNode.Attributes.GetNamedItem("label").Value);
                        if (Get_Child_of_Parent(cNode, false))
                        {
                            _isPrint = true;
                        }
                    }
                    else
                    {
                        if (cNode.HasChildNodes )
                        {
                            //For "Content" nodes having Childrens
                            XmlNodeList conditionnodes = cNode.SelectNodes("Conditions");
                            if (conditionnodes.Count > 0)
                            {
                                if (CheckALLConditions(conditionnodes.Item(0)))
                                {
                                    cNode.Attributes.GetNamedItem("isprint").Value = "true";
                                    if (pNode.Name.ToLower() == "documentsection" && pNode.Attributes.GetNamedItem("preserve").Value == "false")
                                    {
                                        _isPrint = true;
                                    }
                                }
                            }
                            else
                            {  //For Place Holders
                                XmlNodeList PlaceHolders = cNode.SelectNodes("Placeholders");
                                if (PlaceHolders.Count > 0)
                                {
                                    //Handle PlaceHolder in First pass.
                                    //Log.trace("Place Holder "+ PlaceHolders.Count.ToString());
                                }
 
                            }
                        }
                        else
                        {
                            Log.trace(cNode.Name + " getting verified");
                            if (pNode.Name.ToLower() == "documentsections" || pNode.Attributes.GetNamedItem("preserve").Value == "false")
                            {
                                if (Verify_Content(cNode))
                                {
                                    //If it return true; any one contents with is set to true, then set the parent node to true else dont change it.
                                    _isPrint = true;
                                }
                            }
                            else
                            {
                                _isPrint = true;
                            }
                        }
                    }
                }

                if (_isPrint)
                {
                    if (pNode.Attributes.Count > 0)
                    {

                        pNode.Attributes.GetNamedItem("isprint").Value = _isPrint.ToString().ToLower();
                    }
                }
                return _isPrint;
            }


            catch (Exception e)
            {
                MessageBox.Show(e.StackTrace);
                Log.exception(e, e.Message + e.StackTrace);
                return false;
            }

        }

        public bool CheckALLConditions(XmlNode ContentNode)
        {
            bool _condFlag = false;
            XmlNodeList resultNodeList = null;
            string comparator = "";
        

            foreach (XmlNode conditionNode in ContentNode)
            {
                //Get the resulting node list, and then verify if each Node is satisfying the condition.
                resultNodeList = rootICPNode.SelectNodes(conditionNode.Attributes.GetNamedItem("elementPath").Value);
                comparator = conditionNode.Attributes.GetNamedItem("comparator").Value.ToLower();
                _condFlag = false; //reset it everytime

                if (resultNodeList.Count <= 0)
                {
                    //If not nodees are there. Exit out of for loop.
                    if (comparator != "isnull")
                    {
                        return false;
                    }
                }


                switch (comparator)
                {
                    case "isnotnull":
                        foreach (XmlNode resNode in resultNodeList)
                        {
                            if (resNode.InnerXml != null && resNode.InnerXml.Length != 0)
                            { //Check that atleast one of the node has InnerXml with value.
                                //If all are null, then it wont set _condFlag to true;
                                _condFlag = true;
                                break;
                            }
                        }
                        break;

                    case "isnull":

                        
                        //Setting it to true, so if condition fails it will return from method. Else it will continue as true
                        _condFlag = true;

                        foreach (XmlNode resNode in resultNodeList)
                        {
                            if (resNode.InnerXml != null && resNode.InnerXml.Length>0)
                            {
                                _condFlag = false; ///Break out of func
                                return false;
                            }
                        }

                        break;  //if none of the values has value then only it will come over here
                       

                    case "equals":
                        string val = conditionNode.Attributes.GetNamedItem("value").Value.ToLower();
                        foreach (XmlNode resNode in resultNodeList)
                        {
                            if (resNode.InnerXml.ToLower().Equals(val.ToLower()))
                            {
                                _condFlag = true;
                                break;
                            }
                        }
                        break;
                } //End Switch

                if (!_condFlag)
                {
                    return false;  //Return if anyone condition doesnt satisfy.                   
                }

            }//Endfor

            return true;
        }


        /*
         * Method Verify Contents will accept the content node as input
         * checks the type of content and based on it, it check if it has value or its empty.
         * If it has Value for element/TaskList: set Node for Print = true
         * */
        public bool Verify_Content(XmlNode contentNode)
        {
            if (contentNode.Name.ToLower() == "librarycontent")
            {
                //If its library item, it willbe assumed that its going to be Present so always set it to TRUE.
                contentNode.Attributes.GetNamedItem("isprint").Value = "true";
                return true;
            }

            else if (contentNode.Name.ToLower() == "elementcontent")
            {
                //If type is Elements, get path and verify if it has value or not.
                if (DoesElementhasValue(contentNode))
                {
                    contentNode.Attributes.GetNamedItem("isprint").Value = "true";
                    return true;
                }
                else
                {
                    return false;
                }

            }
            else if (contentNode.Name.ToLower() == "taskcontent")
            {
                if (doesTaskListhasValue(contentNode.Attributes.GetNamedItem("taskList").Value))
                {
                    contentNode.Attributes.GetNamedItem("isprint").Value = "true";
                    return true;
                }
                else
                {
                    contentNode.Attributes.GetNamedItem("isprint").Value = "false";
                    return false;
                }
            }
            return false;
        }

        /** Method : Checks if the element has value or not.
         * 
         */
        private bool DoesElementhasValue(XmlNode contentNode)
        {
            XmlNodeList resultNodeList = rootICPNode.SelectNodes(contentNode.Attributes.GetNamedItem("elementPath").Value);


            if (resultNodeList.Count > 0)
            {
                foreach (XmlNode resNode in resultNodeList)
                {
                    if (!(resNode.InnerXml != null && resNode.InnerXml.Length == 0))
                    {
                        return true;
                    }
                }

            }
            //If no value for some reason.
            return false;

        }

        /* Method  "doesTaskListhasValue" will verify if there are any task in specified Tasklist. If not task then it will return false.
         * 
         */

        private bool doesTaskListhasValue(string taskListname)
        {
            bool hasVal = false;


            var tasklist = from task in tl.TaskObjects
                           //where task.Name..Equals("Pinal") & task.Type[1].Equals("Alpha")
                           where (task.Name.Equals(taskListname))
                           select new { objTask = task };


            foreach (var item in tasklist)
            {
                if (item.objTask.ListofTask.Count > 0)
                {
                    hasVal = true;
                }
            }

            if (!hasVal)
            {
                Log.trace(taskListname + " is empty");
            }

            return hasVal;
        }

        #endregion

        #region WriteSection-6

        public bool PrintDocumentSection6(XmlNode docsectionNode, Word.Range wrkRng)
        {
            try
            {
                //Log.trace("Section # : " + WordHelper.getSectionNumber(wrkRng).ToString());
                //string secNum = WordHelper.getSectionNumber(wrkRng).ToString();
                insertmethod = new SectionInsertMethods();
                //insertmethod.wdApp_ = wdApp_;
                //insertmethod.wdDoc = wdDoc_;
                //insertmethod.currdoc_ = tspdDoc_;
                

                Parse_Child_of_Parent(docsectionNode, wrkRng);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                Log.exception(e, e.Message + e.StackTrace);
                return false;
            }
            return true;
        }

        bool preserveFlag = false;
        public void Parse_Child_of_Parent(XmlNode pNode, Word.Range wrkRng)
        {
            // Log.trace("Section # : " + WordHelper.getSectionNumber(wrkRng).ToString());

            if (pNode.Name.ToLower() != "documentsections")
            {
                if (pNode.Attributes.GetNamedItem("isprint").Value == "true")
                {
                    InsertDocumentSection(pNode, wrkRng);
                    Log.trace(pNode.Name + " = Doc Section Inserted");
                }
            }


            foreach (XmlNode cNode in pNode)
            {
                if (cNode.Name.ToLower() == "documentsection" && cNode.Attributes.GetNamedItem("isprint").Value == "true")
                {
                    Parse_Child_of_Parent(cNode, wrkRng);
                }
                else
                {
                    if (cNode.HasChildNodes)
                    {
                        //For "Content" nodes having Childrens

                        //get Preserve attribute for Parent node, so if empty task list we can print message.

                        if (cNode.Attributes.GetNamedItem("isprint").Value == "true")
                        {
                            if (pNode.Name.ToLower() == "documentsection" && pNode.Attributes.GetNamedItem("preserve").Value == "true")
                            {
                                preserveFlag = true;
                            }
                            else
                            {
                                preserveFlag = false;
                            }
                            InsertContent(cNode, wrkRng, preserveFlag);
                        }

                    }
                    else
                    {
                        if (cNode.Attributes.GetNamedItem("isprint").Value == "true")
                        {
                            //get Preserve attribute for Parent node, so if empty task list we can print.
                            if (pNode.Name.ToLower() == "documentsections" || pNode.Attributes.GetNamedItem("preserve").Value == "true")
                            {
                                preserveFlag = true;
                            }
                            else
                            {
                                preserveFlag = false;
                            }
                            InsertContent(cNode, wrkRng, preserveFlag);
                        }
                    }
                } //End Else
            } //EndFOR

            //  }//END IF
        }

        public void InsertDocumentSection(XmlNode sectionNode, Word.Range wrkRng)
        {
            if (sectionNode.Name.ToLower() == "documentsection")
            {
                try
                {
                    tspdDoc_.getStyleHelper().setNamedStyle(sectionNode.Attributes.GetNamedItem("headingStyle").Value, wrkRng);
                }
                catch (Exception ex)
                {
                    // MessageBox.Show(ex.ToString());
                }
                wrkRng.InsertAfter(sectionNode.Attributes.GetNamedItem("label").Value);
                wrkRng.InsertParagraphAfter();
                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
            }

        }

        public bool InsertContent(XmlNode contentNode, Word.Range _selRng, bool preserveFlag)
        {
            if (contentNode.Name.ToLower() == "librarycontent")
            {
                //If its library item, it willbe assumed that its going to be Present so always set it to TRUE.
                InsertLibraryItem(contentNode, _selRng);
                mc.setStyle(mc.getMessageByName("normalstyle").Format.Style, tspdDoc_, _selRng);
                return true;
            }
            else if (contentNode.Name.ToLower() == "textcontent")
            {
                //If its library item, it willbe assumed that its going to be Present so always set it to TRUE.

                string freeText = contentNode.Attributes.GetNamedItem("value").Value;

                if (freeText.Length > 0)
                {
                    setStyle(contentNode.Attributes.GetNamedItem("bodyStyle").Value, _selRng);
                    _selRng.InsertAfter(freeText);
                    _selRng.InsertParagraphAfter();
                    _selRng.Collapse(ref WordHelper.COLLAPSE_END);
                }
                mc.setStyle(mc.getMessageByName("normalstyle").Format.Style, tspdDoc_, _selRng);
                return true;
            }
            else if (contentNode.Name.ToLower() == "elementcontent")
            {
                //If type is Elements, get path and verify if it has value or not.
                if (contentNode.Attributes.GetNamedItem("elementPath").Value.Length > 0)
                {
                    string seqPath = "";
                    if (contentNode.Attributes.GetNamedItem("sortByPath") != null)
                    {
                        seqPath = contentNode.Attributes.GetNamedItem("sortByPath").Value;
                    }

                    bool softReturn = false;
                    if (contentNode.Attributes.GetNamedItem("softReturn") != null)
                    {
                        softReturn = Convert.ToBoolean(contentNode.Attributes.GetNamedItem("softReturn").Value);
                    }

                    InsertContentbyelementPath(contentNode.Attributes.GetNamedItem("elementPath").Value, seqPath, softReturn,contentNode.Attributes.GetNamedItem("bodyStyle").Value, _selRng);
                    mc.setStyle(mc.getMessageByName("normalstyle").Format.Style, tspdDoc_, _selRng);
                }

            }
            else if (contentNode.Name.ToLower() == "taskcontent")
            {
                _selRng.Collapse(ref WordHelper.COLLAPSE_END);
                InsertTaskList(contentNode, _selRng, preserveFlag);
                mc.setStyle(mc.getMessageByName("normalstyle").Format.Style, tspdDoc_, _selRng);
            }
            else if (contentNode.Name.ToLower() == "methodcontent")
            {
                //Getting Table Caption; If Any
                string tablecaption = "";
                if (contentNode.Attributes.GetNamedItem("tablecaption") != null)
                {
                    tablecaption = contentNode.Attributes.GetNamedItem("tablecaption").Value;
                }

                //Getting Show Caption; value if table caption is to be inserted.
                bool showtablecaption = false;
                if (contentNode.Attributes.GetNamedItem("showtablecaption") != null && contentNode.Attributes.GetNamedItem("showtablecaption").Value.Length>0)
                {
                    showtablecaption = Convert.ToBoolean(contentNode.Attributes.GetNamedItem("showtablecaption").Value);
                }

                string outcometype = "";
                if (contentNode.Attributes.GetNamedItem("param") != null)
                {
                    outcometype = contentNode.Attributes.GetNamedItem("param").Value;
                }

              

                InsertbySpecificMethods(contentNode.Attributes.GetNamedItem("name").Value, tablecaption, _selRng, outcometype,showtablecaption);

               mc.setStyle(mc.getMessageByName("normalstyle").Format.Style, tspdDoc_, _selRng);
            }
            return false;
        }

        ArrayList tskList = new ArrayList();
        ArrayList arrListofTask = new ArrayList();

        public void InsertContentbyelementPath(string xpathQuery,string seqQuery, bool softReturn,string style, Word.Range _selRng)
        {
            try
            {
                XmlNodeList nodeList = rootICPNode.SelectNodes(xpathQuery);
                ArrayList seqArraylist = new ArrayList();
                ArrayList arraylistElem = new ArrayList();

                string[] arrayElem;
                int[]  arraySeq;
                int seq,i=0;

                
                //Get Sequence;
                if (seqQuery.Length > 0)
                {
                    XmlNodeList seqNodelist = rootICPNode.SelectNodes(seqQuery);
                    foreach (XmlNode seqnode in seqNodelist)
                    {
                        if (seqnode.InnerText.Length > 0)
                        {
                            seq = Convert.ToInt16(seqnode.InnerText);
                            seqArraylist.Add(seq);
                            i++;
                        }
                    }

                }

             
                    arraySeq = (int[])seqArraylist.ToArray(typeof(int));
              

                i = 0;
               //Get Elements
                foreach (XmlNode cnode in nodeList)
                {
                    if (cnode.InnerText.Length > 0)
                    {
                        //arrayElem[i]=cnode.InnerText;
                        arraylistElem.Add(cnode.InnerText);
                        i++;
                    }
                }

              
                    arrayElem = (string[])arraylistElem.ToArray(typeof(string));
              
               
                if (arrayElem.Length>0 && arrayElem.Length == arraySeq.Length)
                {//If there are some elements to print.

                    Array.Sort(arraySeq, arrayElem);
                }

                string val = "";
                foreach (string str in arrayElem)
                {
                    setStyle(style, _selRng);
                    //_selRng.InsertAfter(cnode.InnerText);
                    if (softReturn)
                    {
                        val = str.Replace("\n", "\v");
                    }
                    else
                    {
                        val = str;
                    }
                    WordFormatter.FTToWordFormat2(ref _selRng, val);
                    _selRng.InsertParagraphAfter();
                    _selRng.Collapse(ref WordHelper.COLLAPSE_END);
 
                }

                ////Run the Xpath query which will return a nodelist. Print each nodes "Inner XML"
                //foreach (XmlNode cnode in nodeList)
                //{
                //    if (cnode.InnerText.Length > 0)
                //    {
                //        setStyle(style, _selRng);
                //        //_selRng.InsertAfter(cnode.InnerText);
                //        Formatter.FTToWordFormat2(_selRng, cnode.InnerText);
                //        _selRng.InsertParagraphAfter();
                //        _selRng.Collapse(ref WordHelper.COLLAPSE_END);
                //    }
                //}
            }
            catch (Exception e)
            {
                Log.exception(e, xpathQuery + " ==> " + e.Message);
                throw e;
            }
        }

        public void setStyle(string styleName, Word.Range _selRng)
        {
            try
            {
                tspdDoc_.getStyleHelper().setNamedStyle(styleName, _selRng);
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.ToString());
            }

        }
        public void InsertLibraryItem(XmlNode libNode, Word.Range _selRng)
        {
            try
            {
                string finaldate="";
                string librichText = "";
                string elementval = "";
                string bucketName = libNode.Attributes.GetNamedItem("category").Value;
                string libitemName = libNode.Attributes.GetNamedItem("libraryItem").Value;
                LibraryElement Libelem = lm.getLibraryElement(bucketName, libitemName);
                //tspdDoc_.insertLibraryItemByName(Libelem.getElementName(), _selRng);


                XmlNodeList PlaceHolders = libNode.SelectNodes("Placeholders");
                if (PlaceHolders.Count > 0)
                {

                    //Code for handling PlaceHolders
                    #region PlaceHolders

                    string path = BridgeProxy.getInstance().loadLibraryElement(Libelem.getLibraryBucketID(), Libelem.getPKValue());
                    Log.trace("Library Item Path:- " + path);
                    if (System.IO.File.Exists(path))
                    {
                        //	richTextBox1.Clear();
                        if (System.IO.Path.GetExtension(path) == ".rtf")
                        {

                            RichTextBox rtflibctl = new RichTextBox();
                            rtflibctl.Visible = false;
                            rtflibctl.LoadFile(path);


                            foreach (XmlNode placeholdernode in PlaceHolders.Item(0))
                            {
                                Log.trace(placeholdernode.Attributes.GetNamedItem("tag").Value + "--->" + placeholdernode.Attributes.GetNamedItem("elementPath").Value);
                                XmlNode valNode = rootICPNode.SelectSingleNode(placeholdernode.Attributes.GetNamedItem("elementPath").Value);
                               
                               
                                //If No Value for an ELEMENT, then do not do anything. Brian will take care it based on condition

                                if (valNode != null && valNode.InnerXml.Length > 0)
                                {
                                    if (placeholdernode.Attributes.GetNamedItem("function") != null && placeholdernode.Attributes.GetNamedItem("function").Value.Length > 0)
                                    {
                                        if (placeholdernode.Attributes.GetNamedItem("function").Value.ToLower() == "converttoqqyyyy")
                                        {
                                            //Converting Dates to Quarter
                                            finaldate = ConvertToQuarter(valNode.InnerText);
                                            rtflibctl.Rtf = rtflibctl.Rtf.Replace(placeholdernode.Attributes.GetNamedItem("tag").Value, finaldate);
                                        }

                                    }
                                    else  ///If No Function, but regular Element replace
                                    {
                                        if (placeholdernode.Attributes.GetNamedItem("codeList") != null && placeholdernode.Attributes.GetNamedItem("codeList").Value.Trim().Length > 0)
                                        {
                                            //IF there is CodeList, get userlabel
                                            elementval = bom_.getIcpSchemaMgr().getUserLabel(placeholdernode.Attributes.GetNamedItem("codeList").Value, valNode.InnerText);
                                        }
                                        else
                                        {
                                            elementval = valNode.InnerText;
                                        }

                                        Log.trace("FTT Formatting TAGS: " + Formatter.stripFormatInstruction(elementval));
                                        //Assign the Rich Text to RTF Control; Also strippped FT Formatting tags.
                                        rtflibctl.Rtf = rtflibctl.Rtf.Replace(placeholdernode.Attributes.GetNamedItem("tag").Value, Formatter.stripFormatInstruction(elementval));
                                    }
                                }                                
                            }

                            rtflibctl.SelectAll();
                            rtflibctl.Copy();

                            //Pasting from RichText Box
                            _selRng.Paste();  //To preserve formatting.
                            _selRng.Collapse(ref WordHelper.COLLAPSE_END);
                            _selRng.InsertParagraphAfter();  //T0 Stop leaking of range in to next document section. 
                            
                        }
                    }
                    #endregion
                }
                else
                {  //Regular Library Item, With NO Placeholders
                    Word.Range tmpRng = tspdDoc_.insertLibraryItemByNameNonInteractive(Libelem.getElementName(), _selRng);

                    tmpRng.Start = tmpRng.Start - 1;
                    if (tmpRng.Text == "\r")  //Remove extra paragraph markers.
                    {
                        tmpRng.Text = tmpRng.Text.Replace("\r", "");
                        // tmpRng.Start = tmpRng.Start;
                        tmpRng.Collapse(ref WordHelper.COLLAPSE_START);
                    }

                    _selRng.Start = tmpRng.Start;
                    _selRng.End = tmpRng.End;

                    _selRng.Collapse(ref WordHelper.COLLAPSE_END);   //Collapsing range, incase.
                }
            }
            catch (Exception ex)
            {
                string msg = mc.getMessageByName("exception3").Text;
                msg = msg.Replace("[[libitemname]]", libNode.Attributes.GetNamedItem("libraryItem").Value);
                _selRng.InsertAfter(msg);
                _selRng.InsertParagraphAfter();
                mc.setStyle(mc.getMessageByName("exception3").Format.Style, tspdDoc_, _selRng);
                _selRng.Collapse(ref WordHelper.COLLAPSE_END);
                Log.exception(ex, ex.Message);
            }
            finally
            {
                _selRng.Collapse(ref WordHelper.COLLAPSE_END);
            }
        }

        private string ConvertToQuarter(string enrollmentdate)
        {
            DateTime ft;
            try
            {
                 ft = Convert.ToDateTime(enrollmentdate);
            }
            catch (Exception ex)
            {
                Log.exception(ex,"Date Conversion Error: " + ex.Message);
                throw ex;
            }
            int year = ft.Year;
            int month = ft.Month;
            string Qyear = "Q";
            if (month >= 1 && month <= 3)
            {
                Qyear += "1";
            }
            else if (month >= 4 && month <= 6)
            {
                Qyear += "2";
            }
            else if (month >=7 && month <= 9)
            {
                Qyear += "3";
            }
            else if (month >= 10 && month <= 12)
            {
                Qyear += "4";
            }

            Qyear += " " + year.ToString();
            return Qyear;
        }

        public void InsertTaskList(XmlNode tskNode, Word.Range _selRng, bool preserveFlag)
        {
            //preseverFlag = TRUE, means a task list is empty and it has to be preserved. Then Print a message.

            string msg = "";
          
            string taskListname = tskNode.Attributes.GetNamedItem("taskList").Value;
            string referencemesg = mc.getMessageByName("reference").Text;
            arrListofTask.Clear();
            arrListofTask = GetTaskList(taskListname);  //Arraylist
            string taskListType = GetEpochforselectedTaskList(taskListname);

            if (arrListofTask.Count <= 0 && preserveFlag)
            {
                msg = tskNode.Attributes.GetNamedItem("message").Value;
                if (msg.Length > 0)
                {
                    _selRng.InsertAfter(msg);
                    _selRng.InsertParagraphAfter();
                    _selRng.Collapse(ref WordHelper.COLLAPSE_END);
                }
                return;  //If no task 
            }
            string stylename = tskNode.Attributes.GetNamedItem("bodyStyle").Value;
            string taskname = "", taskdesc = "", varheader = "", varList = "";
            bool isReferenced = false;
            string sectionLevel= "";
            SectionMappingConfig.message m1 = null;
            bool insertpara = false;

            //Loop thru each task
            foreach (Task tsk in arrListofTask)
            {
                isReferenced = false;
                insertpara = false;
                varList = "";
                //Load the variables with details
                taskname = tsk.getActualDisplayValue();
                sectionLevel ="";
                if (taskListType == "epoch")
                {
                    sectionLevel = Verify_Reference(tsk, taskListname);

                    if (sectionLevel.Length > 0)
                    {
                        taskdesc = referencemesg + sectionLevel;
                        isReferenced = true;
                    }
                    else
                    {
                        taskdesc = tsk.getFullDescription();
                    }
                }
                else
                {
                    taskdesc = tsk.getFullDescription();
                }

                taskdesc = taskdesc.Replace("\n", "\v");

                varheader = "";
                if (!isReferenced)
                {
                    varList = VariablesbyTask(tsk);
                    if (varList.Trim().Length <= 0)
                    {
                        varList = mc.getMessageByName("exception1").Text;
                    }
                }
                //Start writing FirstLine with formatting
                m1 = secConfig.getMessageByName(stylename, "firstline");
                msg = m1.Text;
                msg = msg.Replace("[[task]]", taskname);
                msg = msg.Replace("[[taskdetails]]", taskdesc);
                msg = msg.Replace("[[variablelist]]", varList);
                msg = Formatter.formatCleaner(msg);

                if (!m1.Format.isNewLine)
                {
                    //Set a Flag to "True" if newLine is false. So we can insert paragraph regardless if second line has text or not.
                    insertpara = true;  
                }

                if (msg.Trim().Length > 0)
                {
                    _selRng.InsertAfter(msg);
                    if (m1.Format.isNewLine)
                    {
                        _selRng.InsertParagraphAfter();
                        secConfig.setStyle(m1.Format.Style, tspdDoc_, _selRng);
                        _selRng.Collapse(ref WordHelper.COLLAPSE_END);
                    }
                    else
                    {
                        secConfig.setStyle(m1.Format.Style, tspdDoc_, _selRng);
                        _selRng.Collapse(ref WordHelper.COLLAPSE_END);
                    }
                }

                //Start writing Second Line with formatting
                m1 = secConfig.getMessageByName(stylename, "secondline");
                msg = m1.Text;
                msg = msg.Replace("[[task]]", taskname);
                msg = msg.Replace("[[taskdetails]]", taskdesc);
                msg = msg.Replace("[[variablelist]]", varList);
                msg = Formatter.formatCleaner(msg);
                if (msg.Trim().Length > 0)
                {
                    // secConfig.setStyle(secConfig.getMessageByName(stylename, "secondline").Format.Style, tspdDoc_, _selRng);
                    _selRng.InsertAfter(msg);

                    if (m1.Format.isNewLine)
                    {
                        secConfig.setStyle(m1.Format.Style, tspdDoc_, _selRng);
                        _selRng.InsertParagraphAfter();
                        _selRng.Collapse(ref WordHelper.COLLAPSE_END);
                    }
                    else
                    {
                        secConfig.setStyle(m1.Format.Style, tspdDoc_, _selRng);
                        _selRng.Collapse(ref WordHelper.COLLAPSE_END);
                    }
                }
                else
                {
                    //If not data to print, we need to insert paragraph; else all will merge in 1st line.
                    if (insertpara)
                    {
                        _selRng.InsertParagraphAfter();
                        _selRng.Collapse(ref WordHelper.COLLAPSE_END);
                    }
                }

                if (!isReferenced)
                {
                    //Start writing ThirdLine with formatting
                    m1 = secConfig.getMessageByName(stylename, "thirdline");
                    msg = m1.Text;
                    msg = msg.Replace("[[task]]", taskname);
                    msg = msg.Replace("[[taskdetails]]", taskdesc);
                    msg = msg.Replace("[[variablelist]]", varList);
                    msg = Formatter.formatCleaner(msg);
                    if (msg.Trim().Length > 0)
                    {
                        _selRng.InsertAfter(msg);

                        if (m1.Format.isNewLine)
                        {
                            _selRng.InsertParagraphAfter();
                            secConfig.setStyle(m1.Format.Style, tspdDoc_, _selRng);
                            _selRng.Collapse(ref WordHelper.COLLAPSE_END);
                        }
                        else
                        {
                            secConfig.setStyle(m1.Format.Style, tspdDoc_, _selRng);
                            _selRng.Collapse(ref WordHelper.COLLAPSE_END);
                        }
                    }

                    //Start writing 4th Line with formatting
                    m1 = secConfig.getMessageByName(stylename, "fourthline");
                    msg = m1.Text;
                    msg = msg.Replace("[[task]]", taskname);
                    msg = msg.Replace("[[taskdetails]]", taskdesc);
                    msg = msg.Replace("[[variablelist]]", varList);
                    msg = Formatter.formatCleaner(msg);
                    if (msg.Trim().Length > 0)
                    {
                        //  secConfig.setStyle(secConfig.getMessageByName(stylename, "fourthline").Format.Style, tspdDoc_, _selRng);
                        _selRng.InsertAfter(msg);

                        if (m1.Format.isNewLine)
                        {
                            secConfig.setStyle(m1.Format.Style, tspdDoc_, _selRng);
                            _selRng.InsertParagraphAfter();
                            _selRng.Collapse(ref WordHelper.COLLAPSE_END);
                        }
                        else
                        {
                            secConfig.setStyle(m1.Format.Style, tspdDoc_, _selRng);
                            _selRng.Collapse(ref WordHelper.COLLAPSE_END);
                        }
                    }
                } // is referenced.
            } //END FOR
        }

        public void InsertbySpecificMethods(string MethodName,string TableCaption, Word.Range _selRng,string outcometype,bool showtblCaption)
        {
            switch (MethodName.ToLower())
            {
                //case "insertbloodvolumetable":
                //    insertmethod.InsertBloodVolumeTable(_selRng, mc, tl, TableCaption, showtblCaption);
                //    break;
                //case "insertdosingtable":
                //    insertmethod.InsertDosingTable(_selRng, mc, TableCaption, false, showtblCaption);
                //    break;
                //case "insertadditionaldosingtable":
                //    insertmethod.InsertDosingTable(_selRng, mc, TableCaption, true, showtblCaption);
                //    break;
                //case "insertdispensingtable":
                //    insertmethod.InsertDispensingTable(_selRng, mc, TableCaption, "drugdispensing", showtblCaption);
                //    break;
                //case "insertadditionaldispensingtable":
                //    insertmethod.InsertDispensingTable(_selRng, mc, TableCaption, "AdditionalDrugDispensing", showtblCaption);
                //    break;
                //case "insertinvestigationalproducttable":
                //    insertmethod.InsertInvestigationalProductTable(_selRng, mc, TableCaption, this.icpInstMgr_, showtblCaption);
                //    break;
                //case "insertstudydesigntable":
                //    insertmethod.InsertStudyDesignTable(_selRng, mc, TableCaption, this.icpInstMgr_, showtblCaption);
                //    break;
                //case "insertemergencycontactstable":
                //    insertmethod.InsertEmergencyContacts(_selRng, mc, TableCaption, showtblCaption);
                //    break;
                //case "insertsponsororg":                    
                //    insertmethod.InsertSponsorOrg(_selRng, mc);
                //    break;
                //case "insertleadinvestigator":
                //    insertmethod.InsertLeadInvestigator(_selRng, mc,this.icpInstMgr_);
                //    break;
                ////InsertAnalysisVariableByType
                //case "insertanalysisvariablebytype":
                //    insertmethod.InsertAnalysisVariableByType(_selRng, mc, this.icpInstMgr_, outcometype);
                //    break;
                //case "insertstatsmodel":
                //    insertmethod.InsertStatsModel(_selRng, mc);
                //    break;
                //case "insertpopulationsets":
                //    insertmethod.InsertPopulationSets(_selRng, mc);
                //    break;
                //case "insertanalyses":
                //    insertmethod.InsertAnalyses(_selRng, mc);
                //    break;
            }
 
        }

        private string Verify_Reference(Task tsk,string _currentTaskListname)
        {
            //This method returns the ArrayList of Task for provided taskListname
            //tskList.Clear();
            int index = -1;
            string val="";
            var tasklist = from task in tl.TaskObjects
                           //where task.Name..Equals("Pinal") & task.Type[1].Equals("Alpha")
                           where (task.ListofTask.Contains(tsk) && task.Type.Equals("purpose"))
                           select new { objTask = task };


            foreach (var item in tasklist)
            {
                if (item.objTask.Name != _currentTaskListname)
                {

                    //XmlNode myNode = docsectionNode.SelectSingleNode("//Macros/Macro[@name='Collection of Study Variables']//DocumentSection[Content[@taskList='Efficacy']]/@secNum");
                    XmlNode myNode = docsectionNode.SelectSingleNode("//DocumentSection[TaskContent[@taskList='" + item.objTask.Name + "']]");
                    if (myNode != null && myNode.InnerXml.Length > 0)
                    {
                        Log.trace ("TASK - " + tsk.getActualDisplayValue() + " exists in " + item.objTask.Name + " at position" + item.objTask.ListofTask.IndexOf(tsk).ToString());
                        if (item.objTask.Name.ToLower() == "safety")
                        {
                            int cnt = 0;
                            foreach (XmlNode safetychildnode in myNode.ChildNodes)
                            {
                                if (safetychildnode.Name.ToLower() == "documentsection" && safetychildnode.Attributes.GetNamedItem("isprint").Value =="true")
                                {
                                    cnt++;
                                }
                                else if (safetychildnode.Name.ToLower() == "taskcontent")
                                {
                                    //There might be some task associated with other safety EP;
                                    // so we need get the Indexof; Per Matt's scenario
                                    index = item.objTask.ListofTask.IndexOf(tsk)+ cnt + 1;
                                    break;  
                                }
                            }//end foreach
                        }
                        else
                        {
                            index = item.objTask.ListofTask.IndexOf(tsk) + 1;
                        }

                        val =  myNode.Attributes.GetNamedItem("secNum").Value +  "."+ index.ToString();
                        return val;
                    }
                }
            }
            

            return val;
        }

        private ArrayList GetTaskList(string taskListname)
        {
            //This method returns the ArrayList of Task for provided taskListname
            tskList.Clear();
            var tasklist = from task in tl.TaskObjects
                           //where task.Name..Equals("Pinal") & task.Type[1].Equals("Alpha")
                           where (task.Name.Equals(taskListname))
                           select new { objTask = task };


            foreach (var item in tasklist)
            {
                if (item.objTask.ListofTask.Count > 0)
                {
                    tskList = item.objTask.ListofTask;
                }
            }
            return tskList;
        }

        private string GetEpochforselectedTaskList(string taskListname)
        {
            var tasklist = from task in tl.TaskObjects
                           //where task.Name..Equals("Pinal") & task.Type[1].Equals("Alpha")
                           where (task.Name.Equals(taskListname))
                           select new { objTask = task };


            foreach (var item in tasklist)
            {
                return item.objTask.Type;
            }
            return "";
 
        }

        private string VariablesbyTask(Task sel_Task)
        {
           // ArrayList studyVar = new ArrayList();
            string strVariableList = "";
            IEnumerator mappings = sel_Task.getVariableMappings();
            VariableDictionary dict = bom_.getVariableDictionary();
            StudyVariable selVar = null;
            while (mappings.MoveNext())
            {
                VarRef mapping = (VarRef)mappings.Current;
                selVar = dict.findBySourceID(mapping.getVariableID());
                if (selVar != null)
                {
                  //  studyVar.Add(selVar.getActualDisplayValue() + ", ");
                    strVariableList +=  selVar.getInstructions() + ", ";               
                }
            }//End while

            if (strVariableList.Trim().Length > 0)
            {
                strVariableList = strVariableList.Trim().TrimEnd(',');
            }
            return strVariableList;
        }
        #endregion



        public  void getSectionNumber(XmlNodeList myNodeList, string _parentNumber)
        {
            int cnt = 0;
            foreach (XmlNode secNode in myNodeList)
            {
                if (secNode.Name == "DocumentSection" &&
                    secNode.Attributes.GetNamedItem("isprint").InnerText == "true"
                    && secNode.Attributes.GetNamedItem("level").InnerText != "0")
                {
                    cnt += 1;
                    string sectionNumber = _parentNumber + "." + cnt.ToString();
                    if (secNode.Attributes.GetNamedItem("secNum") != null)
                    {
                        secNode.Attributes.GetNamedItem("secNum").InnerText = sectionNumber;
                        Log.trace(sectionNumber + " " + secNode.Attributes.GetNamedItem("label").InnerText);
                    }
                    if (secNode.HasChildNodes)
                        getSectionNumber(secNode.ChildNodes, sectionNumber);
                }//end if
            }//end foreach
        }//end function



        public override void postProcess()
        {
            // Clean up memory

            Log.trace("Memory before collection:: " + GC.GetTotalMemory(false).ToString());
            _currentSOA = null;
            ParentTaskNode = null; ;
            docsectionNode = null;
            stylenode = null;
            taskListnode = null;
            rootICPNode = null;
            tl = null;
            mc = null;
            insertmethod = null;    
            GC.Collect();
            Log.trace("Memory after collection --> " + GC.GetTotalMemory(true).ToString());
        }
    }
} 


