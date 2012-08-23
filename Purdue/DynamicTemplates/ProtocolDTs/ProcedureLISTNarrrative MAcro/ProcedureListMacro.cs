using System;
using System.Collections;
using System.Xml;
using Tspd.Tspddoc;
using Tspd.MacroBase;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;
using MSXML2;

using System.Windows.Forms;

using Word = Microsoft.Office.Interop.Word;

namespace VersionControl 
{
	internal sealed class ProcedureListMacro
	{
		private static readonly string header_ = @"$Header: ProcedureList.cs, 1, 27-May-2010 10:17:02, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for ProcedureList.
	/// </summary>
	public class ProcedureListMacro : AbstractMacroImpl
	{
		SOA _currentSOA = null;
		long _currentArm = ArmRule.ALL_ARMS;
        bool form_cancelled = false;
        ArrayList _VisitIDwithTaskevents = new ArrayList();
		Task _foundTask = null;
		bool _addTaskHeader = false;
		bool _UsePeriod = false;
        string perORsubper = "";
        long taskID = -1;
        long perID = -1;
        Period selPeriod = null;
        
        MacrosConfig mc = null;

        public ProcedureListMacro(MacroExecutor.MacroParameters mp)
            : base(mp)
		{
		}

		#region Dynamic Tmplt Methods
		
		#region ProcedureList

		public static MacroExecutor.MacroRetCd ProcedureList (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.SalesDemo.DynTmplts.ProcedureListMacro.ProcedureList,ProtocolDTs.dll" elementLabel="SOA Narrative" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Schedule" autogenerates="true" toolTip="TimesByTask" shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>
#endif
            try 
			{
				mp.pba_.setOperation("Procedure List Macro", "Generating information...");

                ProcedureListMacro macro = null;
                macro = new ProcedureListMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Procedure List Macro"); 
				mp.inoutRng_.Text = "Procedure List: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		#endregion 

		#endregion

		public static new bool canRun(BaseProtocolObject bpo)
		{
			SOA soa = bpo as SOA;
			if (soa == null)
			{
				return false;
			}

			if (soa.isSchemaDesignMode()) 
			{
				return false;
			}

			return true;
		}

        public override void preProcess()
        {
            string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);
            if (MacroBaseUtilities.isEmpty(elementPath))
            {
                return;
            }

            try
            {
                //Initiate the configuration file and set your variables
                string chooserElementPath = this.macroEntry_.getElementPath();
                string fPath = tspdDoc_.getTrialProject().getTemplateDirPath() + "\\dyntmplts\\MacrosConfig.xml";
                mc = new MacrosConfig(fPath, chooserElementPath);
            }
            catch (Exception ex)
            {
                Log.exception(ex, ex.Message);
                MessageBox.Show("Configuration file is missing. Please contact your Configuration Administrator", "Procedure List Macro");
            }



            SOAEnumerator soaEnum = bom_.getAllSchedules();
            while (soaEnum.MoveNext())
            {
                pba_.updateProgress(2.0);

                SOA soa = soaEnum.getCurrent();
                if (soa.getElementPath().Equals(elementPath))
                {
                    _currentSOA = soa;
                    break;
                }
            }

            if (_currentSOA == null) return;

            bool isBad = false;


            // Get stored parameters
            string sParms = execParms_.getParm(MacroExecutor.MacroExecParameters.PARM_1);
            string[] aParms = null;

            if (!MacroBaseUtilities.isEmpty(sParms))
            {
                aParms = sParms.Split('|');
            }

            bool parmsValid = true;

            if (aParms != null && aParms.Length == 2)
            {
                if (!MacroBaseUtilities.isEmpty(aParms[0]))
                {
                    perID = PurdueUtil.getNumber(aParms[0], out isBad);
                    if (isBad)
                    {
                        return;
                    }
                }

                if (!MacroBaseUtilities.isEmpty(aParms[1]))
                {
                    try
                    {
                        _UsePeriod = bool.Parse(aParms[1]);
                    }
                    catch (Exception ex)
                    {
                        parmsValid = false;
                    }
                }

            }
            else
            {
                parmsValid = false;
            }


            //Forms Object
             ProcedureList procSelect = new ProcedureList();
             string xmlPath = tspdDoc_.getTrialProject().getTemplateDirPath() + "\\dyntmplts\\ProcDescMapping.xml";
             Period _dummyPeriod = null;

            if (parmsValid)
            {
                //If Valid Parameters, then check if for that period/subperiod has any visit specific events which needs description

                //Assign Period/Sub Period.
                _dummyPeriod = GetPeriodorSubPeriod(perID, _UsePeriod);

                if (_dummyPeriod != null)  //If Period/SubPeriod might have be delayed.
                {
                    if (procSelect.loadItems(_currentSOA, xmlPath, _dummyPeriod, _UsePeriod, true))
                    {
                        parmsValid = false;
                    }
                }
            }


            // Ask the user if the parms are missing/invalid
            if (!parmsValid)
            {
                
                //TaskSelect taskSelect = new TaskSelect();
                
                procSelect.Text = mc.getMessageByName("captiontext").Text;
                procSelect.label1.Text = mc.getMessageByName("label1").Text;
                procSelect.rdPeriod.Text = mc.getMessageByName("rdperiod").Text;
                procSelect.rdSubperiod.Text = mc.getMessageByName("rdsubperiod").Text;
                
                procSelect.loadItems(_currentSOA, xmlPath, _dummyPeriod, _UsePeriod,false);

                DialogResult res = procSelect.ShowDialog();

                if (res == DialogResult.Cancel)
                {
                    form_cancelled = true;
                    pba_.done();
                    this.MacroStatusCode = MacroExecutor.MacroRetCd.Cancelled;
                    return;
                }

                if (res == DialogResult.OK)
                {
                    //procSelect.sel_ObjID 
                    if (!procSelect.sel_ObjID.Equals(-1))
                    {
                        perID = (long)procSelect.sel_ObjID;
                    }
                    else
                    {
                        perID = -1;
                    }
                    //selPeriod = _currentSOA.getPeriodByID(perID);

                    // perORsubPer = "period";
                    if (procSelect.perORsubPer == "period")
                    {
                        _UsePeriod = true;
                    }

                }

                if (!form_cancelled || perID != -1)
                {
                    // save it for next time so we don't ask
                    sParms = perID.ToString() + "|";
                    sParms += _UsePeriod.ToString();
                    execParms_.setParm(MacroExecutor.MacroExecParameters.PARM_1, sParms);
                }
            } //Endif !ParamsValid
           
        }


        private Period GetPeriodorSubPeriod(long id, bool isPeriod)
        {
            Period Per = null;
            if (_UsePeriod)
            {
                Per = _currentSOA.getPeriodByID(perID);
            }
            else
            {
                IList perEnum = _currentSOA.getPeriodEnumerator().getList();
                foreach (Period pr in perEnum)
                {
                    IList sp_List = _currentSOA.getPeriodChildren(pr).getList();
                    {
                        foreach (EventScheduleBase subprd in sp_List)
                        {
                            if (subprd.getObjID().Equals(perID))
                            {
                                Per = (Period)subprd;
                                break;
                            }
                        }
                    }
                }
            }

            return Per;
        }


		public override void display()
		{
			Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;
            string Msg = "";

		//	VariableDictionary dict = bom_.getVariableDictionary();

			pba_.updateProgress(1.0);



            if (!form_cancelled)
            {
                //CODE FOR PLACEHOLDER
            }


			if (_currentSOA == null)
			{
				pba_.updateProgress(70.0);

                Msg = mc.getMessageByName("exception2").Text;
                wrkRng.InsertAfter(Msg);
				//wrkRng.InsertAfter("This schedule that this macro refers to was removed, delete this macro.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);						
				inoutRange.End = wrkRng.End;
				setOutgoingRng(inoutRange);
				wdDoc_.UndoClear();
				return;
			}

            if (form_cancelled)
            {
                pba_.updateProgress(70.0);
                //Do nothing. Just Get out!
                inoutRange.End = wrkRng.End;
                setOutgoingRng(inoutRange);
                wdDoc_.UndoClear();
                return;
 
            }

            if (perID == -1)
            {
                pba_.updateProgress(70.0);
                Msg = mc.getMessageByName("exception1").Text;
                wrkRng.InsertAfter(Msg);
                // wrkRng.InsertAfter("Parameters refer to a task which no longer appears within the schedule.");
                wrkRng.InsertParagraphAfter();
                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                inoutRange.End = wrkRng.End;
                setOutgoingRng(inoutRange);
                wdDoc_.UndoClear();
                return;
 
            }
            

            //CODE FOR GETTING PERIOD.

            Period Per = GetPeriodorSubPeriod(perID, _UsePeriod);

            if (Per == null)
            {
                pba_.updateProgress(70.0);
                Msg = mc.getMessageByName("exception3").Text;
                wrkRng.InsertAfter(Msg);
                // wrkRng.InsertAfter("Parameters refer to a task which no longer appears within the schedule.");
                wrkRng.InsertParagraphAfter();
                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                inoutRange.End = wrkRng.End;
                setOutgoingRng(inoutRange);
                wdDoc_.UndoClear();
                return;
 
            }

            TaskEnumerator tenum = _currentSOA.getTaskEnumerator();
            ArrayList taskList = new ArrayList();

            if (_currentSOA.getProtocolEventCount(Per) > 0)
            {
              IList peEnum =  _currentSOA.getProtocolEventEnumerator(Per).getList();
              foreach (ProtocolEvent ev in peEnum)
              {
                  IList tvEnum = _currentSOA.getTaskVisitsForVisit(ev).getList();
                  foreach (TaskVisit tv in tvEnum)
                  {
                      if (taskList.IndexOf(tv.getAssociatedTaskID()) < 0)
                      {
                          taskList.Add(tv.getAssociatedTaskID());
                      }
                  }
              }
            }

            if (taskList.Count == 0)
            {
                //If no TASK is found for the selected period/subperiod.

                pba_.updateProgress(70.0);
                Msg = mc.getMessageByName("exception4").Text;
                
                if (_UsePeriod)
                {
                    Msg = Msg.Replace("[[seltitle]]","Period");
                }
                else 
                {
                    Msg = Msg.Replace("[[seltitle]]", "Sub-Period");
                }
                Msg = Msg.Replace("[[selection]]",Per.getBriefDescription());

                wrkRng.InsertAfter(Msg);
                
                wrkRng.InsertParagraphAfter();
                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                inoutRange.End = wrkRng.End;
                setOutgoingRng(inoutRange);
                wdDoc_.UndoClear();
                return;
            }
            else
            {
                //if TASK is found, 

                long tskID =-1;
                Task tsk = null;
                bool isBad = false;
                string msg = "";

                try
                {
                    mc.setStyle(mc.getMessageByName("firstline").Format.Style, tspdDoc_, wrkRng);
                    msg = mc.getMessageByName("firstline").Text;
                    msg = msg.Replace("[[selection]]",Per.getBriefDescription());
                    wrkRng.InsertAfter(msg);
                    wrkRng.InsertParagraphAfter();
                    wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                }
                catch (Exception ex)
                { }

                //ArrayList _TaskList;
                //if (taskList.Count > 0)
                //{
                //     _TaskList =ReadXML();  //Fill in the arraylist after reading it the task from xml.
                //}

                foreach (long str in taskList)
                {
                  //  tskID = PurdueUtil.getNumber(str, out isBad);
                    tskID = str;
                    tsk = _currentSOA.getTaskByID(tskID);
                    mc.setStyle(mc.getMessageByName("bodytext").Format.Style, tspdDoc_, wrkRng);
                    msg = mc.getMessageByName("bodytext").Text;
                    msg = msg.Replace("[[task]]",tsk.getBriefDescription());
                    msg = msg.Replace("[[taskdesc]]", tsk.getFullDescription());
                    wrkRng.InsertAfter(msg);
                    wrkRng.InsertParagraphAfter();
                    wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                    Print_TaskEvents(Per, tsk, wrkRng);
                    wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                }
            }

            mc.setStyle(mc.getMessageByName("normalstyle").Format.Style, tspdDoc_, wrkRng);        
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			// Set outgoing range
			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();
		}



        private void Print_TaskEvents(Period _per, Task _task,Word.Range wrkRng)
        {
            string msg = null;
             ProtocolEventEnumerator peEnum = _currentSOA.getProtocolEventEnumerator(_per);
            IList tskEvts = peEnum.getList();
            string teDesc = null;
            if (_currentSOA.getTaskUsageState(_task, _per) != SOA.UsageTriState.None)
            {//Check ONLY if the Task has any events.
                //TaskEvents_ = getTaskVisitsOrderedByVisit(tsk, tskEvts);
                ArrayList tvl = new ArrayList(_currentSOA.getTaskVisitsForTask(_task).getList());

                foreach (ProtocolEvent pe in tskEvts)
                {
                    teDesc = "";
                    for (int i = 0; i < tvl.Count; i++)
                    {                        
                        TaskVisit tv = tvl[i] as TaskVisit;
                        if (tv != null && tv.getAssociatedVisitID() == pe.getObjID())
                        {
                            tv.setViewAngle(TaskVisit.ViewAngle.Task);
                            //note: CHECK THE STANDARDTEXT IN PROCEDURElIST.CS ALSO
                            if (tv.getFullDescription() != null && tv.getFullDescription() != "<No Description is selected.>")
                            {
                                if (!teDesc.Contains(tv.getFullDescription()))
                                {
                                    mc.setStyle(mc.getMessageByName("level2").Format.Style, tspdDoc_, wrkRng);
                                    msg = mc.getMessageByName("level2").Text;
                                    msg = msg.Replace("[[taskeventdesc]]", tv.getFullDescription());
                                    msg = msg.Replace("\n", "\v");
                                    wrkRng.InsertAfter(msg);
                                    wrkRng.InsertParagraphAfter(); 
                                    wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                                }
                            }
                            tvl[i] = null;
                        }
                    }
                }

            }
        }

       

        private ArrayList ReadXML()
        {
            //This method will read xml, and store nodes as objects in an ArrayList.
            ArrayList ListTasks = new ArrayList();
            string xmlPath = tspdDoc_.getTrialProject().getTemplateDirPath() + "\\dyntmplts\\ProcDescMapping.xml";

            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(xmlPath);

                // Select and display all Tasks.
                XmlNodeList nodeList;
                XmlElement root = doc.DocumentElement;
                nodeList = root.SelectNodes("/Tasks");
                foreach (XmlNode taskEntry in nodeList)
                {
                    ListTasks.Add(taskEntry);  //Adding node to ArrayList.
                }
            }
            catch (Exception e)
            {
                Log.exception(e, e.Message + " - Configuration file is missing. Please, contact your configuration administrator");
                return ListTasks;
            }
            return ListTasks;
        }

        private void set_Style(string styleName, Word.Range currRng)
        {
            try
            {
                 tspdDoc_.getStyleHelper().setNamedStyle(styleName, currRng);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

		public override void postProcess()
		{
			// Clean up memory
			_currentSOA = null;
			_currentArm = ArmRule.ALL_ARMS;
			
			_foundTask = null;
			_addTaskHeader = false;
		}
	}
}



//* unUsed code *
  ///*   private void DisplaySeq(Word.Range selRng)
  //      {
  //          int arrayIndex = 0;
  //          ArrayList _fullTaskVisitList = new ArrayList();
  //          ArrayList aiList = new ArrayList();
  //          TaskVisitEnumerator en = new TaskVisitEnumerator(
  //          _currentSOA.getTaskVisitForTaskID(_foundTask.getObjID()),
  //          this.icdSchemaMgr_.getTVTemplate());

  //          foreach (TaskVisit tv in en.getList())
  //          {
  //              try
  //              {
  //                  _fullTaskVisitList.Add(tv);
  //              }
  //              catch (Exception ex)
  //              {
  //                  Log.exception(ex, "Duplicate sequence in Task Sequencing");
  //              }


  //              TaskVisitEnumerator children = _currentSOA.getCopiesOfTaskVisit(tv);
  //              foreach (TaskVisit tvSub in children.getList())
  //              {
  //                  try
  //                  {
  //                      _fullTaskVisitList.Add(tvSub);
  //                  }
  //                  catch (Exception ex)
  //                  {
  //                      Log.exception(ex, "Duplicate sequence in Task Sequencing");
  //                  }
  //              }
  //          }

  //          _fullTaskVisitList.Sort(new BusinessObjectFactory.SequenceSort());
  //          selRng.Collapse(ref WordHelper.COLLAPSE_END);
  //          set_Style(style_Bullet, selRng);
  //          string tempVisitId = "0";
  //          ProtocolEvent tempVisit = null;
  //          int CntUnspec = 0;  //Counter for unspecified.

  //          foreach (TaskVisit tv in _fullTaskVisitList)
  //          {
  //              tempVisit = _currentSOA.getVisitOfTaskVisit(tv);
  //              if (tempVisitId.Equals(tempVisit.getObjID().ToString()))
  //              {
  //                 // continue;
  //              }
  //              else
  //              {
  //                  set_Style(style_Bullet, selRng);
  //                  tempVisitId = tempVisit.getObjID().ToString();
  //                  selRng.InsertAfter(tempVisit.getBriefDescription());
  //                  selRng.InsertParagraphAfter();
  //                  selRng.Collapse(ref WordHelper.COLLAPSE_END);
  //                  set_Style(style_SubsidaryList, selRng);
  //              }

  //              string start = getStartText(tv, arrayIndex) + getDurationString(tv);
  //              aiList.Add(start);              
  //              selRng.InsertAfter(start);
  //              arrayIndex++;
  //              selRng.InsertParagraphAfter();
  //              selRng.Collapse(ref WordHelper.COLLAPSE_END);
  //          }

           

  //          foreach (TaskVisit tv in _fullTaskVisitList)
  //          {
  //              string start = getStartText(tv, arrayIndex) + getDurationString(tv);
  //              aiList.Add(start);

  //              if (start.ToLower() == "unspecified")
  //              {
  //                  CntUnspec++;
  //              }
  //              arrayIndex++;
  //          }

  //          //Delete the ENTRY if there is ONLY ONE instance of "unspecified"
  //          if (CntUnspec == 1)
  //          {
  //              aiList.RemoveAt(aiList.IndexOf("Unspecified"));
  //          }

  //          foreach (string str in aiList)
  //          {
  //              selRng.InsertAfter(str);
  //              selRng.InsertParagraphAfter();
  //              selRng.Collapse(ref WordHelper.COLLAPSE_END);
  //          }
          
  //      }
  //* /