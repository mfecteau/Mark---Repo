using System;
using System.Windows.Forms;
using System.Data;
using System.Collections;
using System.Reflection;
using System.Globalization;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;

namespace SignaturePage
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	class Program
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		/// 

		
		[STAThread]
		static void Main(string[] args)
		{
			try 
			{				
				SignaturePage(args[0],args[1],args[2],args[3]);				
			}
			catch (Exception ex) 
			{
				ErrorForm dlg = new ErrorForm();

				dlg.setErrorMessage(ex.Message + "\r\n" + ex.StackTrace);
				dlg.ShowDialog();
			}
		}

		public Word.Selection sel_ = null;
		public Word.Document myDoc = null;
		Word.Application wdApp = null;

		//Local Variable for creating a logical data table.
		private string stageL1 ="Study Outline: Locked";
		

		public string doc_Creationdate ="", doc_finalizationDate ="";

		public DataTable AuditTable = new DataTable("Audit Table");
        public ArrayList AuditColumns = new ArrayList();

    
		// String constants for querying datasets and storing results in Reading Aliases (if any)

        public string stage1Alias;

		public DataTable myView =  new DataTable();
		public ArrayList lstView = new ArrayList();

		
		public static string sFilename = null;
        public static string path4xml = null;
        public string templateDirPath = null;
        public static string dfilename = null;
        public string verDate = null;
        public string trialAuthor = null;

		public string ProtocolID = null;
		public string	aliasFilePath="";

		private string err_mesg = "Value cannot be obtained. Stages were locked out of sequence";

        public static void SignaturePage(string PID, string authorinfo, string sourcePath,string pathXml) 
		{
            //SourcePath --> Source Path for Signature Page.doc
            //authorinfo= displayName<>Title<>email^;    
            //PID = protocolID



			try
			{
                Program p1 = new Program();
                path4xml = pathXml;
              //  MessageBox.Show(path4xml);
                p1.templateDirPath = sourcePath;
                p1.Load1_NewXMLData();

             //   MessageBox.Show("XML Data is successfully transformed into Data Table");
                if (!p1.LoadVersionDate())
                {
                    return; 
                }
            //    MessageBox.Show("Version Date is found.");
                object falsch = false;
                object truth = true;

				object missing = System.Reflection.Missing.Value;
				

				Word.Application wdApp_temp = new Word.Application();

                Copyfile(sourcePath + "\\SignaturePage.doc");
               // p1.templateDirPath = sourcePath; //to read MetricsConfig.XML file
                object filePath = dfilename;
                
				object oFileName = "test";
				object password = "";

                //Word.Document wdDoc1 = wdApp_temp.Documents.Add(ref missing, ref missing, ref missing, ref oTrue);
                //wdApp_temp.Visible = true;
                //object f1 = "Protocol_ABC_SignaturePage";
                Word.Document wdDoc1 = null;    

				try
				{
                   wdDoc1 = wdApp_temp.Documents.Open(ref filePath, ref falsch, ref falsch, ref falsch, ref password,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref truth, ref missing, ref missing, ref missing, ref missing);
                    wdApp_temp.Visible = true;
                    //wdDoc1.SaveAs(ref f1, ref missing, ref missing, ref missing,
                    //    ref missing, ref missing, ref missing, ref missing,
                    //    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
				}
				catch(Exception ex)
				{
					MessageBox.Show("Only one instance of report is permitted.","Signature Page Report",MessageBoxButtons.OK,MessageBoxIcon.Information);
					return;
				}


				object oStart = wdDoc1.ActiveWindow.Selection.Range.Start;
             //   MessageBox.Show("Initiating Writing...");
                p1.sel_ = wdDoc1.ActiveWindow.Selection;
                p1.myDoc = wdDoc1;
                p1.wdApp = wdApp_temp;
                p1.ProtocolID = PID;
                p1.trialAuthor = authorinfo;
                ////p1.trial_Ind = indication;
                p1.buildDocument();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}

        public static void Copyfile(string sourcePath)
        {
            string sfilename;       
            sfilename = sourcePath; 
            dfilename = Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\" + System.IO.Path.GetFileName(sfilename);
            dfilename = System.IO.Path.GetDirectoryName(dfilename) + "\\" + System.IO.Path.GetFileNameWithoutExtension(dfilename) + "_" + "SignaturePage" + ".doc";

            if (System.IO.File.Exists(dfilename) == true)
            {
                //System.IO.File.Delete(dfilename);
                try
                {
                    System.IO.File.Copy(sfilename, dfilename, true);
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("Only one instance of report is permitted.", "ICD Report");
                    //					dfilename = System.IO.Path.GetDirectoryName(dfilename) + "\\" + System.IO.Path.GetFileNameWithoutExtension(dfilename) + "_" + getIcpMgr().getDisplayValue(AdminDefines.ProtocolID, "") + "_" + (i+1) +".doc";
                    //					System.IO.File.Copy(sfilename,dfilename,true);
                }
            }
            else
            {
                System.IO.File.Copy(sfilename, dfilename, true);
            }
        }

		public void buildDocument()
		{	// no generic header.

            this.addReportHeader();
            this.addReportBody();
            this.addReportFooter();
           		
		}

		public  void addReportHeader() 
		{
            enterHeaderFooter(Word.WdSeekView.wdSeekCurrentPageHeader);
            myReplaceALL("<Protocol ID>", ProtocolID);
            exitHeaderFooter();
		}

		
		public  void addReportFooter() 
		{
            enterHeaderFooter(Word.WdSeekView.wdSeekCurrentPageFooter);
            myReplaceALL("<Version Date>", verDate);
            exitHeaderFooter();
		}

		public void exitHeaderFooter()
		{
			myDoc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
			myDoc.ActiveWindow.DocumentMap = false;
		}
        
		public  void addReportBody()
		{        
            Load_Authors();
            myDoc.UndoClear();           
		}

        private void Load_Authors()
        {
            object wdCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            string displayName = "";
            string email = "";
            string title = "";
            string dept = "";
            string company = "";
            bool otherEmail = false;
           string[] authors = trialAuthor.Split('^');
           if (authors.Length > 0)
           {
               Word.Range Rng_ = sel_.Range;
               int tblstartRng = 0,tblendRng =0;
               IEnumerator allTables_ = myDoc.Tables.GetEnumerator();

               while (allTables_.MoveNext())
               {
                   Word.Table tbl = (Word.Table)allTables_.Current;
                   tblstartRng = tbl.Range.Start;
                   tblendRng = tbl.Range.End +1;
                   break;  //exit after first table.
               }

               Word.Range myRng = sel_.Range.Duplicate;
               myRng.SetRange(tblstartRng, tblendRng);
               myRng.Copy();  //copying the Authors table

               Word.Range currRng = myRng.Duplicate;
               //currRng.Collapse(ref wdCollapseEnd);

               int recCnt = 0;
               
               foreach (string author in authors)
               {
                   if (author.Trim().Length > 0)
                   {
                       if (recCnt > 0)
                       {
                           currRng.InsertParagraph();
                           currRng.Paste();
                       }

                       recCnt++;
                       myDoc.ActiveWindow.Selection.SetRange(currRng.Start, currRng.End);
                       Word.Selection mySel = myDoc.ActiveWindow.Selection;
                       currRng.Collapse(ref wdCollapseEnd);

                       title = "";
                       dept = "";
                       company = "";

                       //Parse the string to seperate DisplayName <> Title <> Email
                       string[] info = author.Split('~');
                       displayName = info[0];
                       title = info[1];

                       //Splitting the Title, Company and Department 
                       string[] temp = title.Split('|');

                       int len = temp.Length;
                 
                       if (len == 0)
                       {
                           title = "";
                           dept = "";
                           company = "";
                          // MessageBox.Show("Here");
                       }
                       else if (len ==1)
                       {
                           if (temp[0] != "null")
                           {
                               title = temp[0];
                           }
                           else
                           {
                               title = ""; 
                           }
                       }
                       else if (len == 2)
                       {
                           title = temp[0];
                           dept = temp[1];
                       }
                       else if (len == 3)
                       {
                           title = temp[0];
                           dept = temp[1];
                           company = temp[2];
                       }
                       else
                       {
                           //MessageBox.Show(len.ToString());
                       }
                       
                         
                       //Email
                       email = info[2];
                       if (email.Length > 0)
                       {
                           if (!email.ToLower().Contains("@pharma.com"))
                           {
                               otherEmail = true;
                           }
                       }
                      int newEndRng = myReplaceALL(mySel, "<Display Name>", displayName);

                      myDoc.ActiveWindow.Selection.SetRange(currRng.Start, newEndRng);
                      mySel = myDoc.ActiveWindow.Selection;

                      newEndRng= myReplaceALL(mySel, "<Title>", title);
                      myDoc.ActiveWindow.Selection.SetRange(currRng.Start, newEndRng);
                      mySel = myDoc.ActiveWindow.Selection;

                      newEndRng =  myReplaceALL(mySel, "<Department>", dept);
                       myDoc.ActiveWindow.Selection.SetRange(currRng.Start, newEndRng);
                       mySel = myDoc.ActiveWindow.Selection;

                       
                       newEndRng = myReplaceALL(mySel, "<Company>", company);
                       myDoc.ActiveWindow.Selection.SetRange(currRng.Start, newEndRng);
                       mySel = myDoc.ActiveWindow.Selection;

                   }   
               } //EndFOR

               //Remove the last Approver table if not needed.
               //if otherEmail = TRUE 
               if (!otherEmail)
               {
                   try
                   {
                        myDoc.Tables[myDoc.Tables.Count].Delete();
                   }
                   catch (Exception ex)
                   {
                       MessageBox.Show(ex.ToString() + ": Error deleting approver's table");
                   }
               }

               sel_.Collapse(ref wdCollapseEnd);
           }
        }

        public int myReplaceALL(Word.Selection sel_,string findstring, string replaceString)
        {
            try
            {
                object _optMissing = System.Reflection.Missing.Value;
                object replace = Word.WdReplace.wdReplaceOne;
                
               // Word.Range sel_ = reportDoc_.ActiveWindow.Selection;
                sel_.Find.Replacement.ClearFormatting();
                sel_.Find.Text = findstring;
                sel_.Find.Replacement.Text = replaceString;
                sel_.Find.Forward = false;
                sel_.Find.Wrap = Word.WdFindWrap.wdFindContinue;
                sel_.Find.MatchCase = false;
                sel_.Find.MatchWholeWord = false;

                sel_.Find.Execute(
                    ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing,
                    ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing,
                    ref _optMissing, ref _optMissing, ref replace,
                    ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing);

                return sel_.End;
            }
            catch (Exception ex)
            {
            }
            return sel_.Range.End;
        }

        public void myReplaceALL(string findstring, string replaceString)
        {
        //    MessageBox.Show(replaceString);
            try
            {
                bool found = false;
                object _optMissing = System.Reflection.Missing.Value;
                object replace = Word.WdReplace.wdReplaceAll;
                Word.Selection sel_ = myDoc.ActiveWindow.Selection;
                sel_.Find.Replacement.ClearFormatting();
                sel_.Find.Text = findstring;
                sel_.Find.Replacement.Text = replaceString;
                sel_.Find.Forward = true;
                sel_.Find.Wrap = Word.WdFindWrap.wdFindContinue;
                sel_.Find.MatchCase = false;
                sel_.Find.MatchWholeWord = false;

              found= sel_.Find.Execute(
                    ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing,
                    ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing,
                    ref _optMissing, ref _optMissing, ref replace,
                    ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing);
         //     MessageBox.Show("Replacing in Header: " + found.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

		private bool LoadVersionDate()
		{
			
            string finalversion = "Final Version";
            verDate = geteventDate(finalversion);
            if (verDate == null)
            {
             //If version date is null, try to look if study design concepts is locked, if not then you cannot run this report.
                if (ReadAliases2(": Locked"))
                {
                    if (stage1Alias.Length > 0)
                    {
                        verDate = geteventDate(stage1Alias);
                    }
                    if (verDate == null)
                    {
                        MessageBox.Show("Study Design elements must first be locked before this report can be generated.", "Signature Page Report", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return false;
                    }
                }
            }           
            return true;
		}

		private int StringtoInt(string days)
		{
			try
			{
				int temp =  Convert.ToInt16(days);
				return temp;
			}
			catch(Exception ex)
			{
				return 0;
			}			
		}

        private bool ReadAliases2(string state)
        {
            //This function returns a string (expression) for the query to dataset.
            //state  == Locked or Unlocked.
            //Modifies the 

            //aliasFilePath = templateDirPath.ToString() +"/TSPDTemplates/" + trial_ID + "/dyntmplts/MetricsConfig.xml";		
            aliasFilePath = templateDirPath + "\\MetricsConfig.xml";
            stage1Alias = "";

            if (System.IO.File.Exists(aliasFilePath) == true)
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(aliasFilePath);
                // Select and display all Audit Entries.
                XmlNodeList nodeList;
                XmlElement root = doc.DocumentElement;
                // nodeList = root.SelectNodes("/MetricsReport/ConceptGroups/ConceptGroup");
                nodeList = doc.GetElementsByTagName("ConceptGroup");
                foreach (XmlNode AliasEntry in nodeList)
                {
                    if (AliasEntry.Attributes[0].InnerText.ToLower() == "studydesign")
                    {
                        stage1Alias = "COMMENTS LIKE " + "'*" + AliasEntry.Attributes[1].InnerText + state + "*'";
                        foreach (XmlNode xiNode in AliasEntry.ChildNodes)
                        {
                            if (xiNode.InnerText.Length > 0)
                            {
                                stage1Alias += " OR COMMENTS LIKE " + " '*" + xiNode.InnerText.ToString() + state + "*' ";
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Configuration File is missing. Please contact your configration administrator", "Signature Page Report", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            return true;
        }


		private string geteventDate(string actionname)
		{
           // string strExpr1 = "COMMENTS LIKE" + "'*" + actionname + "*'";

      //      MessageBox.Show("Looking for Modify Date");
            string strExpr1 = actionname;
            string selectCommand1 = "MAX (MODIFY_DATE)";
            try
            {
                object s1 = (object)AuditTable.Compute(selectCommand1, strExpr1);
                DateTime d1 = (DateTime)s1;
                return d1.ToString("dd MMM yyyy hh:mm:ss tt");
            }
            catch (Exception ex)
            {
          //      MessageBox.Show("Modify Date not found");
                return null;
            }
		}


        private string StageLockDate(string stagename1)
        {
            //Returns the MOST recent date at which the stage was locked.
            string strExpr1 = stagename1;
            string selectCommand1 = "MAX (MODIFY_DATE)";
            try
            {
                object s1 = (object)AuditTable.Compute(selectCommand1, strExpr1);
                DateTime d1 = (DateTime)s1;
                return d1.ToString("dd MMM yyyy hh:mm:ss tt");
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private void Load1_NewXMLData()
        {
            int i = 0;
          //  sFilename = "C:\\AuditExport.xml";

            XmlDocument doc = new XmlDocument();
            doc.Load(path4xml);

            // Select and display all Audit Entries.
            XmlNodeList nodeList;
            XmlElement root = doc.DocumentElement;
            ///nodeList = root.SelectNodes("/audit_hist_collection/audit_hist");
            ///
            nodeList = doc.GetElementsByTagName("AUDIT_HIST");
            //root.SelectNodes("/AUDIT_HIST");

            Load_Columns(nodeList.Item(0));
            int colIdx = 0;

            foreach (XmlNode AuditEntry in nodeList)
            {
                DataRow dtAuditRow = AuditTable.NewRow();

                string strDt = "";

                if (AuditEntry.HasChildNodes)
                {
                    colIdx = 0;
                    foreach (XmlNode xiNode in AuditEntry.ChildNodes)
                    {
                        if (xiNode.Name.ToLower().Contains("date"))
                        {
                            string strdate = xiNode.InnerText.Replace("\r\n", "");
                            if (strdate.Trim().Length > 0)
                            {
                                DateTime d1 = DateTime.ParseExact(strdate, "MM/d/yyyy H:mm:ss zzz", DateTimeFormatInfo.InvariantInfo);
                                strDt = d1.ToString("dd MMM yyyy hh:mm:ss tt");
                                dtAuditRow[colIdx] = strDt;
                            }
                            else
                            {
                                dtAuditRow[colIdx] = Convert.DBNull;
                            }
                        }
                        else
                        {
                            dtAuditRow[colIdx] = xiNode.InnerText.Replace("\r\n", "");
                        }
                        colIdx++;
                    }
                    AuditTable.Rows.Add(dtAuditRow);
                  
                }
            } //end for		
        }

        private void Load_Columns(XmlNode AuditEntry)
        {
            AuditTable.Clear();  //Clearing the table structure
            System.Type nodeType = null;
            foreach (XmlNode xiNode in AuditEntry.ChildNodes)
            {
                if (xiNode.Name.ToLower().Contains("date"))
                {
                    nodeType = System.Type.GetType("System.DateTime");
                }
                else
                {
                    nodeType = System.Type.GetType("System.String");
                }

                DataColumn col = new DataColumn(xiNode.Name, nodeType);
                AuditTable.Columns.Add(col);
                AuditColumns.Add(xiNode.Name);
            }
        }

	  protected void enterHeaderFooter(Word.WdSeekView where)
		{
			if (!(myDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)) 
			{
				myDoc.ActiveWindow.Panes[2].Close();
			}
			Word.View view = myDoc.ActiveWindow.ActivePane.View;
			if(view.Type == Word.WdViewType.wdNormalView
				|| view.Type == Word.WdViewType.wdOutlineView 
				|| view.Type == Word.WdViewType.wdMasterView )
			{
				view.Type = Word.WdViewType.wdPrintView;
			}
			view.SeekView = where; 
		}
	}
}
