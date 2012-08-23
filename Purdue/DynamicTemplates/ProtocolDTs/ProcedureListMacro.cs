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
        long taskID = -1;
        long perID = -1;
        Period selPeriod = null;
        string selectedTYPE = "";
        ArrayList seltaskListforVisit = new ArrayList();
        bool includeAlltask = false;
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
            string[] storedtaskID = null;
                
            if (!MacroBaseUtilities.isEmpty(sParms))
            {
                aParms = sParms.Split('|');
            }

            bool parmsValid = true; 

            #region handlingOldParameters
            //Due to recent change in param's, this code converts the old param's into new params.
            if (aParms != null && aParms.Length == 2)
            {
                if (!MacroBaseUtilities.isEmpty(aParms[0]))
                {
                    perID = PurdueUtil.getNumber(aParms[0], out isBad);
                    if (isBad)
                    {
                        return;
                    }
                    sParms = aParms[0] + "|";
                }

                if (!MacroBaseUtilities.isEmpty(aParms[1]))
                {
                    try
                    {
                        if (bool.Parse(aParms[1]))
                        {
                            aParms[1] = "period";
                            sParms += aParms[1] + "|";        
                        }
                        else
                        {
                            aParms[1] = "subperiod";
                            sParms += aParms[1] + "|"; 
                        }

                        
                    }
                    catch (Exception ex)
                    {
                        //parmsValid = false;
                    }
                }
            }

            #endregion

            if (!MacroBaseUtilities.isEmpty(sParms))
            {
                aParms = sParms.Split('|');
            }


            if (aParms != null && aParms.Length == 3)
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
                        selectedTYPE = aParms[1];
                    }
                    catch (Exception ex)
                    {
                        parmsValid = false;
                    }
                }
                if (!MacroBaseUtilities.isEmpty(aParms[2]))
                {
                    try
                    {   //Retriving all stored Task ID's

                        if (aParms[2] != "True")
                        {
                            storedtaskID = aParms[2].Split('>');
                            if (storedtaskID.Length > 0)
                            {
                              //  long id = 0;
                                foreach (string _tskID in storedtaskID)
                                {
                                    //id = PurdueUtil.getNumber(_tskID, out isBad);
                                    if (_tskID.Length> 0)
                                    {
                                        seltaskListforVisit.Add(_tskID);
                                    }
                                }
                            }
                        }
                        else
                        {
                            includeAlltask = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        parmsValid = false;
                    }
                }
                else
                {
                    includeAlltask = true;
                }
            }
            else
            {
                parmsValid = false;
            }


            //Forms Object
             ProcedureList procSelect = new ProcedureList();      
             Period _dummyPeriod = null;
             ProtocolEvent _visit = null;

             if (parmsValid)
             {
                 //If Valid Parameters, then check if for that period/subperiod has any visit specific events which needs description
                 if (selectedTYPE != "visit")
                 {
                     //Assign Period/Sub Period.
                     _dummyPeriod = GetPeriodorSubPeriod(perID);
                 }
                 else
                 {
                     _visit = _currentSOA.getProtocolEventByID(perID);                    
                 }


                 if (!includeAlltask)
                 {
                     int cnt = 0;
                     ArrayList removeTaskwithnoEvents = new ArrayList();
                     long tskid = 0;

                     if (selectedTYPE == "visit")
                     {
                         //check if task exists or taskevent exists, 2 checks with one method. If not then remove from parameter list.
                         foreach (string str in seltaskListforVisit)
                         {
                             tskid = Convert.ToInt64(str);
                             if (!_currentSOA.hasTaskVisit(tskid, perID))
                             {
                                 removeTaskwithnoEvents.Add(cnt);
                             }
                             cnt++;
                         }

                         //Removing from parameter list, all the taskids with no events.
                         cnt = 0;
                         foreach (int idx in removeTaskwithnoEvents)
                         {
                             seltaskListforVisit.RemoveAt(idx - cnt);
                             cnt++;
                         }
                     }
                     else   //For period & Subperiod
                     {
                         //check if task exists or taskevent exists, 2 checks with one method. If not then remove from parameter list.
                         foreach (string str in seltaskListforVisit)
                         {
                             tskid = Convert.ToInt64(str);
                             Task t = _currentSOA.getTaskByID(tskid);
                             
                             if (t == null ||  _currentSOA.getTaskUsageState(t, _dummyPeriod) == SOA.UsageTriState.None)
                             {
                                 removeTaskwithnoEvents.Add(cnt);
                             }
                             cnt++;
                         }

                         //Removing from parameter list, all the taskids with no events.
                         cnt = 0;
                         foreach (int idx in removeTaskwithnoEvents)
                         {
                             seltaskListforVisit.RemoveAt(idx - cnt);
                             cnt++;
                         }
                     }

                 }


             }


            // Ask the user if the parms are missing/invalid
            if (!parmsValid)
            {

                if (_currentSOA.getPeriodEnumerator().getList().Count <= 0)
                {
                    MessageBox.Show(mc.getMessageByName("exception6").Text);
                    form_cancelled = true;
                    pba_.done();
                    this.MacroStatusCode = MacroExecutor.MacroRetCd.Cancelled;
                    return;
                }


                procSelect.Text = mc.getMessageByName("captiontext").Text;
                procSelect.label1.Text = mc.getMessageByName("label1").Text;
                procSelect.rdPeriod.Text = mc.getMessageByName("rdperiod").Text;
                procSelect.rdSubperiod.Text = mc.getMessageByName("rdsubperiod").Text;
                procSelect.rdVisit.Text = mc.getMessageByName("rdvisit").Text;
                
                procSelect.loadItems(_currentSOA,mc);

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

                    selectedTYPE = procSelect.perORsubPer;

                    includeAlltask = procSelect.chkAlltasks.Checked;

                    //if (selectedTYPE == "visit")
                    //{
                        seltaskListforVisit = procSelect.taskList;  ///IF visit selection, only SELECTED TASKS to be PRITNED
                  //  }

                }

                if (!form_cancelled || perID != -1)
                {
                    // save it for next time so we don't ask
                    sParms = perID.ToString() + "|";
                    sParms += selectedTYPE + "|";

                    if (!includeAlltask)
                    {                   
                        foreach (string tskID in seltaskListforVisit)
                        {
                            if (tskID != "0")
                            {
                                sParms += tskID + ">";
                            }
                        } //Endfor
                    }
                    else
                    {
                        sParms += includeAlltask;  //If ALL Tasks are to be inserted.
                    }
                    execParms_.setParm(MacroExecutor.MacroExecParameters.PARM_1, sParms);
                }
            } //Endif !ParamsValid
        
           
        }


        private Period GetPeriodorSubPeriod(long id)
        {
            Period Per = null;
            if (selectedTYPE == "period")
            {
                Per = _currentSOA.getPeriodByID(perID);
            }
            else if (selectedTYPE == "subperiod")
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

            Period Per = null;
            ProtocolEvent visit = null;
            //CODE FOR GETTING PERIOD.
            if (selectedTYPE != "visit")
            {
                Per = GetPeriodorSubPeriod(perID);
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
            }
            else
            {
                visit = _currentSOA.getProtocolEventByID(perID);
                if (visit == null)
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
            }
            

            TaskEnumerator tenum = _currentSOA.getTaskEnumerator();
            ArrayList taskList = new ArrayList();
       

            if (includeAlltask)
            {
                if (selectedTYPE != "visit")
                {
                    seltaskListforVisit.Clear();

                    if (_currentSOA.getProtocolEventCount(Per) > 0)
                    {
                        IList peEnum = _currentSOA.getProtocolEventEnumerator(Per).getList();
                        foreach (ProtocolEvent ev in peEnum)
                        {
                            IList tvEnum = _currentSOA.getTaskVisitsForVisit(ev).getList();
                            foreach (TaskVisit tv in tvEnum)
                            {
                                if (taskList.IndexOf(tv.getAssociatedTaskID().ToString()) < 0)
                                {
                                    taskList.Add(tv.getAssociatedTaskID().ToString());
                                }
                            }
                        }
                    }
                }
                else
                {

                    IList tvEnum = _currentSOA.getTaskVisitsForVisit(visit).getList();
                    foreach (TaskVisit tv in tvEnum)
                    {
                        if (taskList.IndexOf(tv.getAssociatedTaskID().ToString()) < 0)
                        {
                            taskList.Add(tv.getAssociatedTaskID().ToString());
                        }
                    }

                }
            }
            else
            {
                //If All Tasks are not selected, then use the selected list.
                taskList = seltaskListforVisit;
            }




           

            if (taskList.Count == 0)
            {
                //If no TASK is found for the selected period/subperiod.

                pba_.updateProgress(70.0);
                Msg = mc.getMessageByName("exception4").Text;

                if (selectedTYPE == "period")
                {
                    Msg = Msg.Replace("[[seltitle]]","Period");
                    Msg = Msg.Replace("[[selection]]", Per.getBriefDescription());
                }
                else if (selectedTYPE == "subperiod")
                {
                    Msg = Msg.Replace("[[seltitle]]", "Sub-Period");
                    Msg = Msg.Replace("[[selection]]", Per.getBriefDescription());
                }
                else if (selectedTYPE == "visit")
                {
                    if (_currentSOA.getTaskVisitsForVisit(visit).getList().Count <= 0)
                    {
                        //IF NO TASK VISIT AT ALL.
                        Msg = Msg.Replace("[[seltitle]]", "Visit");
                        Msg = Msg.Replace("[[selection]]", visit.getBriefDescription());
                    }
                    else
                    {
                        //For VISIT: If none of selected task (Storedin parameter) has taskevents for visit the print appropriate mesage.
                        Msg = mc.getMessageByName("exception5").Text;
                        Msg = Msg.Replace("[[seltitle]]", "Visit");
                        Msg = Msg.Replace("[[selection]]", visit.getBriefDescription());
                    }
                }

                //Set normal style.
                mc.setStyle(mc.getMessageByName("normalstyle").Format.Style, tspdDoc_, wrkRng);
                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

                WordFormatter.FTToWordFormat2(ref wrkRng, Msg);               
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

                    if (selectedTYPE != "visit")
                    {
                        msg = msg.Replace("[[selection]]", Per.getBriefDescription());
                    }
                    else
                    {
                        msg = msg.Replace("[[selection]]", visit.getBriefDescription());
                    }
                    if (msg.Length > 0)
                    {
                        WordFormatter.FTToWordFormat2(ref wrkRng, msg);
                        wrkRng.InsertParagraphAfter();
                        wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                    }
                }
                catch (Exception ex)
                { }


                bool flagRestartNo = false;
                string ftnumberstyle = mc.getMessageByName("bodytext").Format.FTNumberStyle;
                string ftbulletstyle = mc.getMessageByName("bodytext").Format.FtBulletStyle;
                string strnormStyle = mc.getMessageByName("bodytext").Format.Style;

                Word.Style _bulletStyle = PurdueUtil.getStyle(WordDoc, ftbulletstyle);
                Word.Style _numberStyle = PurdueUtil.getStyle(WordDoc, ftnumberstyle);
                Word.Style normStyle = PurdueUtil.getStyle(WordDoc, strnormStyle);

                foreach (string str in taskList)
                {
                  //  tskID = PurdueUtil.getNumber(str, out isBad);
                    tskID =  Convert.ToInt64(str);
                    tsk = _currentSOA.getTaskByID(tskID);

                    //Check if selected Task has more taskVisit for selected PER/SubPer/Visit                   


                    if (tsk != null)
                    {
                        mc.setStyle(strnormStyle, tspdDoc_, wrkRng);
                        msg = mc.getMessageByName("bodytext").Text;
                        msg = msg.Replace("[[task]]", tsk.getBriefDescription());
                        msg = msg.Replace("[[taskdesc]]", tsk.getFullDescription());
                        if (msg.Length > 0)
                        {
                            WordFormatter.FTToWordFormat2(ref wrkRng, msg.Trim(), normStyle, _bulletStyle, _numberStyle);
                            if (!flagRestartNo)
                            {
                                PurdueUtil.resartListNumber(wrkRng);
                                flagRestartNo = true;
                            }
                            
                            wrkRng.InsertParagraphAfter();
                            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                        }
                        if (selectedTYPE != "visit")
                        {
                            Print_TaskEvents(Per, tsk, wrkRng);
                        }
                        else
                        {
                            Print_TaskEventsforVisit(visit, tsk, wrkRng);

                        }
                        wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                    }
                }
            }

            mc.setStyle(mc.getMessageByName("normalstyle").Format.Style, tspdDoc_, wrkRng);        
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			// Set outgoing range
			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();
		}

        private void Print_TaskEventsforVisit(ProtocolEvent _visit, Task _task, Word.Range wrkRng)
        {
            string msg = null;
            string teDesc = null;
            ArrayList tvl = new ArrayList(_currentSOA.getTaskVisitsForTask(_task).getList());
            teDesc = "";
            bool flagRestartNo = false;
            string ftnumberstyle = mc.getMessageByName("level2").Format.FTNumberStyle;
            string ftbulletstyle = mc.getMessageByName("level2").Format.FtBulletStyle;
            string strnormStyle = mc.getMessageByName("level2").Format.Style;

            Word.Style _bulletStyle = PurdueUtil.getStyle(WordDoc, ftbulletstyle);
            Word.Style _numberStyle = PurdueUtil.getStyle(WordDoc, ftnumberstyle);
            Word.Style normStyle = PurdueUtil.getStyle(WordDoc, strnormStyle);
            for (int i = 0; i < tvl.Count; i++)
            {
                TaskVisit tv = tvl[i] as TaskVisit;
                if (tv != null && tv.getAssociatedVisitID() == _visit.getObjID())
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
                            msg = msg.Replace("\r", "\v");
                            if (msg.Length > 0)
                            {
                                WordFormatter.FTToWordFormat2(ref wrkRng, msg.Trim(), normStyle, _bulletStyle, _numberStyle);
                                if (!flagRestartNo)
                                {
                                    PurdueUtil.resartListNumber(wrkRng);
                                    flagRestartNo = true;
                                }
                               // wrkRng.InsertAfter(msg);
                                wrkRng.InsertParagraphAfter();
                                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                            }
                        }
                    }
                    tvl[i] = null;
                }
            }
 
        }

        private void Print_TaskEvents(Period _per, Task _task,Word.Range wrkRng)
        {
            string msg = null;
             ProtocolEventEnumerator peEnum = _currentSOA.getProtocolEventEnumerator(_per);
            IList tskEvts = peEnum.getList();
            string teDesc = null;
            bool flagRestartNo = false;
            string ftnumberstyle = mc.getMessageByName("level2").Format.FTNumberStyle;
            string ftbulletstyle = mc.getMessageByName("level2").Format.FtBulletStyle;
            string strnormStyle = mc.getMessageByName("level2").Format.Style;

            Word.Style _bulletStyle = PurdueUtil.getStyle(WordDoc, ftbulletstyle);
            Word.Style _numberStyle = PurdueUtil.getStyle(WordDoc, ftnumberstyle);
            Word.Style normStyle = PurdueUtil.getStyle(WordDoc, strnormStyle);

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
                                    msg = msg.Replace("\r", "\v");
                                    WordFormatter.FTToWordFormat2(ref wrkRng, msg.Trim(), normStyle, _bulletStyle, _numberStyle);
                                    if (!flagRestartNo)
                                    {
                                        PurdueUtil.resartListNumber(wrkRng);
                                        flagRestartNo = true;
                                    }

                                    wrkRng.InsertParagraphAfter(); 
                                    wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                                }
                            }
                            tvl[i] = null;
                        }
                    }
                }//endfor

            }
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