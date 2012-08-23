using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

using Tspd.Tspddoc;
using Tspd.MacroBase;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;
using Tspd.MacroBase.Table;
using TspdCfg.Purdue.DynTmplts.Table;

namespace TspdCfg.Purdue.DynTmplts
{
    public class SectionInsertMethods
    {

        //        #region BloodvolumeTable

        //public Word.Application wdApp_ = null;
        //public Word.Document wdDoc = null;
        //public TspdDocument currdoc_ = null;
        //private object oOneCount = 1;
        //private object oTwoCount = 2;
        //private BusinessObjectMgr bom_ = null;
        //private IcdSchemaManager _icdSchemaMgr = null;
        //private ArrayList priorityList;


        //        public  void InsertBloodVolumeTable(Word.Range wrkRng,MacrosConfig mc,TaskListClass tl,string tablecaption,bool showtablecaption)
        //        {
        //            MacrosConfig.message msg1 = null;
        //            Word.Row newRow = null;
        //            priorityList = new ArrayList();

        //            bom_ = currdoc_.getBom();
        //            _icdSchemaMgr = currdoc_.getIcdSchemaMgr();


        //            //Init Process & Verification for expception
        //            if (CheckforTaskEventwithNoPurpose(wrkRng))
        //            {

        //                return;
        //            }



        //            //PreProcess 
        //           CleanupTaskListclass(tl);


        //            //Check for continuous events, if any Print error message


        //            //Create table.
        //            wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
        //            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

        //            Word.Table tbl = createTable(wrkRng, 2, 5, tablecaption, mc, showtablecaption);
        //            msg1 = mc.getMessageByName("colheader");

        //            //HeaderRow 1
        //            tbl.Rows[1].Cells[1].Range.Text = "Volume of blood to be drawn from each subject";
        //            tbl.Rows[1].Cells.Merge();
        //            mc.setStyle(msg1.Format.Style, currdoc_, tbl.Rows[1].Range);

        //            tbl.Rows[2].Cells[3].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
        //            tbl.Rows[2].Cells[4].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
        //            tbl.Rows[2].Cells[5].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

        //            //HeaderRow 2
        //            tbl.Rows[2].Cells[1].Range.Text = "Assessment";
        //            tbl.Rows[2].Cells[2].Range.Text = "";
        //            tbl.Rows[2].Cells[3].Range.Text = "Sample volume(mL)";
        //            tbl.Rows[2].Cells[4].Range.Text = "No. of samples";
        //            tbl.Rows[2].Cells[5].Range.Text = "Total volume(mL)";
        //            tbl.Rows[2].Cells[1].Merge(tbl.Rows[2].Cells[2]);
        //            mc.setStyle(msg1.Format.Style, currdoc_, tbl.Rows[2].Range);

        //            tbl.Rows[1].HeadingFormat = VBAHelper.iTRUE;
        //           tbl.Rows[2].HeadingFormat = VBAHelper.iTRUE;

        //            msg1 = mc.getMessageByName("rowheader");
        //            MacrosConfig.message msg2 = mc.getMessageByName("datarow");
        //            //Data Row
        //            double bloodVol = 0;
        //            int cnt = 0;
        //            double totalVol = 0;
        //            double grandVol = 0;
        //            int taskeventCnt = 0;

        //            foreach (var item in tl.TaskObjects)
        //            {
        //                foreach (Task tsk in item.ListofTask)
        //                {
        //                    //Getting TaskEvent count  for TASK.                    
        //                    taskeventCnt = GetTaskEventCountforTask(tsk,item.value);

        //                    if (taskeventCnt != 0)
        //                    {
        //                        newRow = tbl.Rows.Add();
        //                        newRow.HeadingFormat = VBAHelper.iFALSE;
        //                        mc.setStyle(msg2.Format.Style, currdoc_, newRow.Range);
        //                        if (newRow.Cells.Count < 5)
        //                        {
        //                            SplittoTwoCell(newRow);
        //                        }
        //                        if (cnt == 0)
        //                        {
        //                            newRow.Cells[1].Range.Text = item.Name;
        //                            mc.setStyle(msg1.Format.Style, currdoc_, newRow.Cells[1].Range);
        //                        }

        //                        newRow.Cells[2].Range.Text = tsk.getActualDisplayValue();
        //                        bloodVol = Convert.ToDouble(tsk.getValueForNode("BloodVolumeml1288637307703").ToString());
        //                        newRow.Cells[3].Range.Text = bloodVol.ToString();
        //                        newRow.Cells[4].Range.Text = taskeventCnt.ToString();
        //                        totalVol = bloodVol * taskeventCnt;
        //                        grandVol += totalVol;
        //                        newRow.Cells[5].Range.Text = totalVol.ToString(); 
        //                        cnt++; //increment count after first printed row, as first cell do no need to be printed all time.
        //                    }                    
        //                }               
        //                cnt = 0;  //Reset it for next tasklist item.
        //            }

        //            newRow = tbl.Rows.Add();
        //            if (newRow.Cells.Count < 5)
        //            {
        //                SplittoTwoCell(newRow);
        //            }
        //            newRow.Cells[1].Range.Text = "Total";
        //            newRow.Cells[5].Range.Text = grandVol.ToString();

        //            wrkRng.End = tbl.Range.End;
        //            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //            wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
        //            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

        //        }


        //        public  bool CheckforTaskEventwithNoPurpose(Word.Range wrkRng)
        //        {

        //            string exception1 = "There is one or more tasks that have a blood volume specified but have an task event without a purpose."; //for missing purpose 

        //            string exception2 = "There is one or more tasks that have a blood volume specified but have been specified as continuous tasks."; //for cont events

        //            string exception3 = "There is one or more tasks that have a blood volume specified but have no task event."; //for No task events


        //            ArrayList taskwithcontEvents = new ArrayList();


        //            foreach (SOA soa in bom_.getAllSchedules().getList())
        //            {

        //                #region HandlingContinuousEvents
        //                //Code below will go thru all Cont events and verify its not the selected task list.
        //                // If a Task with blood vol has cont event, then display error message.

        //                foreach (ProtocolEvent cevt in soa.getStripedEvents())
        //                {
        //                    foreach (TaskVisit tvcont in soa.getTaskVisitsForVisit(cevt).getList())
        //                    {
        //                        if (!taskwithcontEvents.Contains(tvcont.getAssociatedTaskID()))
        //                        {
        //                           taskwithcontEvents.Add(tvcont.getAssociatedTaskID());  //Add TaskID to arrayList
        //                        }
        //                    }
        //                }
        //                #endregion
        //                bool flagHasTaskvisit = false;
        //                IList tskList = soa.getTaskEnumerator().getList();
        //                foreach (Task tsk in tskList)
        //                {
        //                    //Only List Tasks with Blood Volume Filled
        //                    if (tsk.getValueForNode("BloodVolumeml1288637307703") != null && tsk.getValueForNode("BloodVolumeml1288637307703").ToString().Length > 0)
        //                    {

        //                        #region HandlingContinuousEvents
        //                        //Code below will go thru all Cont events and verify its not the selected task list.
        //                        // If a Task with blood vol has cont event, then display error message.


        //                                if (taskwithcontEvents.Contains(tsk.getObjID()))
        //                                {
        //                                    wrkRng.InsertAfter(exception2);
        //                                    wrkRng.InsertParagraphAfter();
        //                                    wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                                    return true;
        //                                }                            

        //                        #endregion

        //                        TaskVisitEnumerator en = new TaskVisitEnumerator(soa.getTaskVisitForTaskID(tsk.getObjID()), _icdSchemaMgr.getTVTemplate());
        //                        IEnumerable <TaskVisit> taskVisits = new List<TaskVisit>();

        //                     //   taskVisits = (List<TaskVisit>)en.getList().Cast<TaskVisit>();
        //                        taskVisits = en.getList().Cast<TaskVisit>();


        //                        if (en.getList().Count > 0)
        //                        {
        //                            flagHasTaskvisit = true;  //Set it to true, if any Task has Taskvisit.
        //                        } 



        //                        //First Pass - NO Purpose at all. OR (we need to verify, because of "Other" Purpose Type - See Design Guide Col on Right side. )
        //                        //Scenario: A Task Visit has only "Other purpose type", taskv.HasPurpose will return true.
        //                        //this DCO need to get Taskvisits associated with outcome types only.

        //                        var taskvisitlist = from taskv in taskVisits
        //                                            where (!taskv.hasPurposes())
        //                                            select new { objTaskvisit = taskv };

        //                        foreach (var item in taskvisitlist)
        //                        {

        //                            wrkRng.InsertAfter(exception1);
        //                            wrkRng.InsertParagraphAfter();
        //                            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                            Log.trace("Task : " + tsk.getActualDisplayValue() + " has visit " + item.objTaskvisit.getActualDisplayValue() + " - No Purpose though Blood Vol defined");                            
        //                            return true;
        //                        } //end Foreach



        //                        //Second Pass - In this pass, we need to filter all the purpose types and make sure they are not the "Other Purpose Type"
        //                         taskvisitlist = from taskv in taskVisits
        //                                            where (taskv.hasPurposes())
        //                                            select new { objTaskvisit = taskv };

        //                        foreach (var item in taskvisitlist)
        //                        {
        //                            bool found = false;
        //                            IEnumerator tvPurposes = soa.getTaskVisitPurposes(item.objTaskvisit);
        //                            while (tvPurposes.MoveNext())
        //                            {
        //                                TaskVisitPurpose tvp = (TaskVisitPurpose)tvPurposes.Current;
        //                                if (!LittleUtilities.isEmpty(tvp.pathToAssociatedOutcome()))
        //                                {
        //                                    found = true;  //See if there is any one instance for Task Visit purpose with outcome.
        //                                    break;
        //                                }
        //                            }

        //                            if (!found)
        //                            {   //If not found then 
        //                                wrkRng.InsertAfter(exception1);
        //                                wrkRng.InsertParagraphAfter();
        //                                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                                Log.trace("Task : " + tsk.getActualDisplayValue() + " has visit " + item.objTaskvisit.getActualDisplayValue() + " - No Purpose though Blood Vol defined");
        //                                return true;
        //                            }
        //                        } //End Foreach

        //                    } //End if task has BLood vol
        //                }//End for each task

        //                if (!flagHasTaskvisit)
        //                {
        //                    //If not found then 
        //                    wrkRng.InsertAfter(exception3);
        //                    wrkRng.InsertParagraphAfter();
        //                    wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                    Log.trace("SOA: " + soa.getActualDisplayValue() + " has task with Bloodvol but not  Task visits "); 
        //                    return true;
        //                }
        //            }
        //            return false;
        //        }

        //        public  int GetTaskEventCountforTask(Task tsk,ArrayList listTypes)
        //        {
        ////          SOAEnumerator soEnum = currdoc_.getBom().getAllSchedules();
        //            int sopCount = 0;
        //            int cnt = 0;
        //            int purposeCnt =0;
        //            bool purposeFound = false;

        //            SOA _SelSOA = bom_.getSchedule(tsk.getScheduleID());
        //            TaskVisitEnumerator en = new TaskVisitEnumerator(_SelSOA.getTaskVisitForTaskID(tsk.getObjID()),_icdSchemaMgr.getTVTemplate());

        //          foreach (TaskVisit tv in en.getList())
        //            {                
        //                if (tv.hasPurposes())
        //                {
        //                    IEnumerator tvPurpose = _SelSOA.getTaskVisitPurposes(tv);
        //                    purposeCnt = 0;
        //                    purposeFound = false;

        //                    while (tvPurpose.MoveNext())
        //                    {
        //                        TaskVisitPurpose tvp = (TaskVisitPurpose)tvPurpose.Current;
        //                        if (!LittleUtilities.isEmpty(tvp.pathToAssociatedOutcome()))
        //                        {
        //                          Outcome o1 = bom_.getOutcomeByPath(tvp.pathToAssociatedOutcome());
        //                            if (o1 != null)
        //                            {
        //                                if (listTypes.IndexOf(o1.getOutcomeType())>=0 )
        //                                {
        //                                    purposeFound = true;
        //                                }
        //                                purposeCnt++;
        //                            }
        //                        }
        //                    }

        //                    if (purposeFound && purposeCnt == 1)
        //                    {
        //                        ProtocolEvent pe = _SelSOA.getProtocolEventByID(tv.getAssociatedVisitID());
        //                        if (pe != null)
        //                        {
        //                            bool status = displayVisitSOP(bom_, _SelSOA, pe, false, false);
        //                            if (status)
        //                            {
        //                                Log.trace(tsk.getActualDisplayValue() + " at Visit: " + pe.getActualDisplayValue() + " has SOP - " + status);

        //                                sopCount = getVisitSOP(bom_, _SelSOA, pe, false, false, tv);
        //                                Log.trace("It has " + sopCount.ToString());
        //                                if (sopCount == 0)
        //                                {
        //                                    cnt++;  //If has SOP, but no checks are done for that task.
        //                                }
        //                                else
        //                                {
        //                                    cnt = cnt + sopCount; 
        //                                }
        //                            }
        //                            else
        //                            {
        //                                cnt++;
        //                            }                            
        //                        }


        //                        // Priority 
        //                       priorityList.Add(tv.getObjID());
        //                        //cnt++;
        //                    }
        //                    else if (purposeFound && !priorityList.Contains(tv.getObjID()))
        //                    {
        //                        priorityList.Add(tv.getObjID());
        //                         ProtocolEvent pe = _SelSOA.getProtocolEventByID(tv.getAssociatedVisitID());
        //                         if (pe != null)
        //                         {
        //                             bool status = displayVisitSOP(bom_, _SelSOA, pe, false, false);
        //                             if (status)
        //                             {
        //                                 sopCount = getVisitSOP(bom_, _SelSOA, pe, false, false, tv);
        //                                 if(sopCount == 0)
        //                                 {
        //                                     cnt++;  //Incase a task has SOP but no timeoiints are checked.
        //                                 }
        //                                 else
        //                                 {
        //                                     cnt = cnt + sopCount;
        //                                 }
        //                             }
        //                             else
        //                             {
        //                                 cnt++;  //If not SOP, then its regular task visit.
        //                             }
        //                         }
        //                         else
        //                         {
        //                             cnt++;
        //                         }
        //                    }

        //                } //End if has purpose
        //            }

        //            return cnt; 
        //        }

        //        public  void SplittoTwoCell(Word.Row _row)
        //        {
        //            //Split the first cell in to two.
        //            _row.Cells[1].Split(ref oOneCount, ref oTwoCount);
        //        }
        //        private  bool CleanupTaskListclass(TaskListClass tl)
        //        {
        //            //This method will remove all the task where Blood Volume is empty.

        //            bool hasTaskwithPurpose = false;

        //            ArrayList indexList= new ArrayList();
        //            int cnt =0;  //for storing index
        //            foreach (var item in tl.TaskObjects)
        //            {
        //                cnt = 0;
        //                foreach (Task tsk in item.ListofTask)
        //                {

        //                    if (tsk.getValueForNode("BloodVolumeml1288637307703") == null || tsk.getValueForNode("BloodVolumeml1288637307703").ToString().Length <= 0)
        //                    {
        //                        indexList.Add(cnt);  //Store index of Task with empty Blood volume 
        //                    }
        //                    else
        //                    {
        //                        hasTaskwithPurpose = true;
        //                    }
        //                    cnt++;
        //                }

        //                //remove all the task who doesnt have blood volume
        //                cnt = 0;
        //                foreach (int idx in indexList)
        //                {

        //                    item.ListofTask.RemoveAt(idx-cnt);
        //                    cnt++; //Have a counter to see how many tasks are removed as each time one would task item would be less.
        //                }

        //                //Clear index for next task list
        //                indexList.Clear();
        //            }

        //            return false;
        //        }

        //        //Handling SOP/Timepoint for each visit

        //        private bool displayVisitSOP(BusinessObjectMgr bom, SOA soa, ProtocolEvent pe, 
        //            bool includeUnusedTasks, bool createMissingEndPointBuckets)
        //        {
        //            bool hasVisitProcedures;

        //            // Get the visit sop matrix
        //            object[,] matrix = ScheduleTree.getVisitProcedures(bom, soa, pe,
        //                includeUnusedTasks, createMissingEndPointBuckets, out hasVisitProcedures);

        //            if (!hasVisitProcedures)
        //            {
        //                return false;
        //            }

        //            int nRows = matrix.GetLength(1);
        //            int nCols = matrix.GetLength(0);

        //            if (nRows > 0)
        //            {
        //                return true;
        //            }

        //            return false;  //Return false
        //        }

        //        private int getVisitSOP(BusinessObjectMgr bom, SOA soa, ProtocolEvent pe,
        //         bool includeUnusedTasks, bool createMissingEndPointBuckets,TaskVisit _currtv)
        //        {
        //            bool hasVisitProcedures;

        //            // Get the visit sop matrix
        //            object[,] matrix = ScheduleTree.getVisitProcedures(bom, soa, pe,
        //                includeUnusedTasks, createMissingEndPointBuckets, out hasVisitProcedures);

        //            if (!hasVisitProcedures)
        //            {
        //                Log.trace("No SOP " + "-1");
        //                return 0;
        //            }

        //            int nRows = matrix.GetLength(1);
        //            int nCols = matrix.GetLength(0);

        //            if (nRows > 0)
        //            {
        //                int cnt=0;
        //                 // Write tv cells, note matrix index starts at 1
        //                for (int j = 1; j < nRows; j++)
        //                {
        //                    // note matrix index starts at 1
        //                    for (int i = 1; i < nCols; i++)
        //                    {
        //                        TaskVisit tv1 = matrix[i, j] as TaskVisit;
        //                        if (tv1 != null && (tv1.getParentID() == _currtv.getObjID() || tv1.getObjID() == _currtv.getObjID()))
        //                        {
        //                            cnt++;
        //                            Log.trace(tv1.getActualDisplayValue() + " - " + tv1.getObjID() + " Count: " + cnt.ToString());
        //                        }
        //                    }
        //                }

        //                return cnt;
        //            }

        //            return 0;  //Return false
        //        }


        //        #endregion

        //        #region Dosingtable

        //        public  void InsertDosingTable(Word.Range wrkRng, MacrosConfig mc,string tblCaption,bool isAdditionalStudyDrug,bool showtablecaption)
        //        {
        //            //this method will print table, note that it will also be used to print table with studydrug whose primary role =AdditionalStudyDrug
        //            //Same code will be called for printing table except "StudyDrug without PrimaryRole="AdditionalStudyDrug"

        //            MacrosConfig.message msg1 = null;
        //            Word.Row newRow = null;

        //            bom_ = currdoc_.getBom();
        //            _icdSchemaMgr = currdoc_.getIcdSchemaMgr();

        //            IList ctmList = bom_.getCTMaterialEnumerator().getList();
        //            IList armEnum = null;
        //            string firstdose = null;
        //            string doseunit="";
        //            string studydrug = "";
        //            string studydrug_form = "";
        //            string route = "";
        //            string primaryrole = "";

        //          //  IList arrTypes = bom_.getIcpSchemaMgr().getEnumPairs("WeightTypes");
        //          //EnumPair ep = 
        //            msg1 = mc.getMessageByName("colheader");
        //            MacrosConfig.message msg2 = mc.getMessageByName("datarow");

        //            foreach (SOA _currentSOA in bom_.getAllSchedules().getList())
        //            {
        //                if (armEnum != null)
        //                {
        //                    armEnum.Clear();
        //                }
        //                armEnum = bom_.getArmsForAssociatedSchedule(_currentSOA).getList();

        //                if (armEnum.Count > 0)
        //                {
        //                    //Create table.
        //                    wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
        //                    wrkRng.Collapse(ref WordHelper.COLLAPSE_END);


        //                    #region SettingTableCaption
        //                    string captiontext = tblCaption;
        //                    //Custom caption if more the one SOA
        //                    if (bom_.getAllSchedules().getList().Count > 1)
        //                    {
        //                        if (isAdditionalStudyDrug)
        //                        {
        //                            captiontext = tblCaption +  " (" + _currentSOA.getActualDisplayValue() + " )";
        //                        }
        //                        else
        //                        {
        //                            captiontext = tblCaption + " (" + _currentSOA.getActualDisplayValue() + " )";
        //                        }
        //                    }

        //                    #endregion


        //                    Word.Table tbl = createTable(wrkRng, 1, 7, captiontext,mc,showtablecaption);

        //                    //First Col Header
        //                    #region firstcolumnheader
        //                    tbl.Rows[1].Cells[1].Range.Text = "Treatment Group";
        //                    tbl.Rows[1].Cells[2].Range.Text = "Investigational Product";
        //                    tbl.Rows[1].Cells[3].Range.Text = "Form";
        //                    tbl.Rows[1].Cells[4].Range.Text = "Route";
        //                    tbl.Rows[1].Cells[5].Range.Text = "Dose";
        //                    tbl.Rows[1].Cells[6].Range.Text = "Composition";
        //                    tbl.Rows[1].Cells[7].Range.Text = "Frequency";
        //                    //tbl.Rows[1].Cells[8].Range.Text = "First Dose";
        //                    //tbl.Rows[1].Cells[9].Range.Text = "Duration";

        //                    tbl.Rows[1].HeadingFormat = VBAHelper.iTRUE;
        //                    currdoc_.getActiveWordDocument().UndoClear();
        //                    #endregion

        //                    //DataRows
        //                    #region dataRow

        //                    bool _printRow = false;

        //                    foreach (Arm _arm in armEnum)
        //                    {
        //                        IList associationList = bom_.getCTMaterialToArmForArm(_arm).getList();

        //                        newRow = tbl.Rows.Add();
        //                        newRow.HeadingFormat = VBAHelper.iFALSE;
        //                        //Cell 1 only to be filled.
        //                        newRow.Cells[1].Range.Text = _arm.getActualDisplayValue();

        //                        //Check if associationList.count>0 Check to see if one of the association is "Additional or not" depeind


        //                        if (associationList.Count > 0)
        //                        {
        //                            foreach (CTMaterialToArm ctaToarm in associationList)
        //                            {  
        //                                ClinicalTrialMaterial ctm = bom_.getCTMaterial(ctaToarm.getAssociatedMaterialID());

        //                                //Print all except CTM with Role = "AdditionalStudyDrug"
        //                                if (ctm.getParentID() != 0)
        //                                {
        //                                    ClinicalTrialMaterial ctmParent = bom_.getCTMaterial(ctm.getParentID());
        //                                    primaryrole = ctmParent.getPrimaryRole();                                  
        //                                }
        //                                else
        //                                {
        //                                    primaryrole = ctm.getPrimaryRole();  //If CTM is Parent.
        //                                }


        //                                if (ctm != null)
        //                                {
        //                                    _printRow = false;

        //                                        if (isAdditionalStudyDrug && primaryrole.ToLower() == "additionalstudydrug")
        //                                        {
        //                                            //flagset
        //                                            _printRow = true;
        //                                        }
        //                                        else if (!isAdditionalStudyDrug && primaryrole.ToLower() != "additionalstudydrug")
        //                                        {
        //                                            // flagset 
        //                                            _printRow = true;
        //                                        }

        //                                        if (_printRow)
        //                                        {

        //                                            newRow = tbl.Rows.Add();

        //                                            if (ctm.getParentID() != 0)
        //                                            {
        //                                                ClinicalTrialMaterial ctmParent = bom_.getCTMaterial(ctm.getParentID());
        //                                                studydrug = ctmParent.getActualDisplayValue();

        //                                                //Formulation
        //                                                if (ctmParent.getFormulation() != null && ctmParent.getFormulation().ToLower() == "other")
        //                                                {
        //                                                    studydrug_form = ctmParent.getOtherFormulation();
        //                                                }
        //                                                else
        //                                                {

        //                                                    studydrug_form = bom_.getIcpSchemaMgr().getUserLabel("FormulationTypes", ctmParent.getFormulation());   
        //                                                   // studydrug_form = ctmParent.getFormulation();
        //                                                }

        //                                                //Route
        //                                                if (ctmParent.getRoute() != null && ctmParent.getRoute().ToLower() == "other")
        //                                                {
        //                                                    route = ctmParent.getOtherRoute();
        //                                                }
        //                                                else
        //                                                {
        //                                                    route = bom_.getIcpSchemaMgr().getUserLabel("RouteOfAdminTypes", ctmParent.getRoute());   
        //                                                    //route = ctmParent.getRoute();
        //                                                }
        //                                            }
        //                                            else //If CTM is Parent
        //                                            {
        //                                                studydrug = ctm.getActualDisplayValue();

        //                                                //Formulation
        //                                                if (ctm.getFormulation() != null && ctm.getFormulation().ToLower() == "other")
        //                                                {
        //                                                    studydrug_form = ctm.getOtherFormulation();
        //                                                }
        //                                                else
        //                                                {
        //                                                    studydrug_form = bom_.getIcpSchemaMgr().getUserLabel("FormulationTypes", ctm.getFormulation());   
        //                                                }


        //                                                //Route
        //                                                if (ctm.getRoute() != null && ctm.getRoute().ToLower() == "other")
        //                                                {
        //                                                    route = ctm.getOtherRoute();
        //                                                }
        //                                                else
        //                                                {
        //                                                    route = bom_.getIcpSchemaMgr().getUserLabel("RouteOfAdminTypes", ctm.getRoute());   
        //                                                }
        //                                            }
        //                                            //newRow.Cells[2].Range.Text = studydrug;
        //                                            Formatter.FTToWordFormat2(newRow.Cells[2].Range, studydrug);
        //                                            newRow.Cells[3].Range.Text = studydrug_form;
        //                                            newRow.Cells[4].Range.Text = route;

        //                                            if (primaryrole != "placebo")
        //                                            {
        //                                                if (ctm.getDoseUnit() != null && ctm.getDoseUnit().ToLower() == "other")
        //                                                {
        //                                                    doseunit = ctm.getDose() + ctm.getOtherDoseUnit();                                                
        //                                                }
        //                                                else
        //                                                {
        //                                                    doseunit = ctm.getDose() + bom_.getIcpSchemaMgr().getUserLabel("WeightTypes", ctm.getDoseUnit());
        //                                                }
        //                                            }
        //                                            else
        //                                            {  //If Role =  Placebo, do not print dose. Just print Placebo
        //                                                doseunit = "placebo";
        //                                            }

        //                                            if (doseunit != null && doseunit.Length > 0)
        //                                            {
        //                                                Formatter.FTToWordFormat2(newRow.Cells[5].Range, doseunit);
        //                                            }
        //                                          //  newRow.Cells[5].Range.Text = doseunit;

        //                                            if (ctm.getStrength() != null && ctm.getStrength().Length>0)
        //                                            {
        //                                                Formatter.FTToWordFormat2(newRow.Cells[6].Range, ctm.getStrength());
        //                                            }
        //                                           // newRow.Cells[6].Range.Text = ctm.getStrength();
        //                                            newRow.Cells[7].Range.Text = bom_.getIcpSchemaMgr().getUserLabel("TimeIntervalTypes", ctm.getFrequencyUnit());

        //                                            mc.setStyle(msg2.Format.Style, currdoc_, tbl.Range); //First apply style to whole table
        //                                            #region commentedCode

        //                                            //firstdose = "";
        //                                            //firstdose = GetFirstTaskVisit(_currentSOA, ctm.getMaterialName(), ctm.getDose(), doseunit);
        //                                            //newRow.Cells[8].Range.Text = firstdose;
        //                                            //newRow.Cells[9].Range.Text = "duratioN";

        //                                            //Add rows for each DOSE
        //                                            ////foreach (ClinicalTrialMaterial ctmChild in ctm.getChildrenLikeParent().getList())
        //                                            ////   {
        //                                            ////       newRow = tbl.Rows.Add();
        //                                            ////       newRow.Cells[2].Range.Text = ctmChild.getMaterialName();
        //                                            ////       if (ctmChild.getFormulation() != null && ctmChild.getFormulation().ToLower() == "other")
        //                                            ////       {
        //                                            ////           newRow.Cells[3].Range.Text = ctmChild.getOtherFormulation();
        //                                            ////       }
        //                                            ////       else
        //                                            ////       {
        //                                            ////           newRow.Cells[3].Range.Text = ctmChild.getFormulation();
        //                                            ////       }
        //                                            ////       if (ctmChild.getRoute() != null && ctmChild.getRoute().ToLower() == "other")
        //                                            ////       {
        //                                            ////           newRow.Cells[4].Range.Text = ctmChild.getOtherRoute();
        //                                            ////       }
        //                                            ////       else
        //                                            ////       {
        //                                            ////           newRow.Cells[4].Range.Text = ctmChild.getRoute();
        //                                            ////       }
        //                                            ////       newRow.Cells[5].Range.Text = ctmChild.getDose() + doseunit;
        //                                            ////       newRow.Cells[6].Range.Text = ctmChild.getStrength();
        //                                            ////       newRow.Cells[7].Range.Text = bom_.getIcpSchemaMgr().getUserLabel("TimeIntervalTypes", ctm.getFrequencyUnit());

        //                                            ////       //newRow.Cells[8].Range.Text = firstdose;
        //                                            ////       //newRow.Cells[9].Range.Text = "duratioN";

        //                                            ////   }//End Child
        //                                            #endregion
        //                                        } //End _printRow

        //                                }
        //                            } //End Foreach CTMaterial
        //                        }
        //                        else
        //                        {
        //                            //If no Association for a study arm, please print message. 
        //                            newRow = tbl.Rows.Add();
        //                            newRow.Cells[1].Range.Text = "No association to Study drug for this Arm";
        //                            newRow.Cells.Merge();
        //                            mc.setStyle(msg2.Format.Style, currdoc_, newRow.Range);
        //                        }
        //                    }

        //                   mc.setStyle(msg1.Format.Style, currdoc_, tbl.Rows[1].Range);  //Header style for First Row.
        //                   //Log.trace("STYLE: " + msg1.Name + "- " + msg1.Format.Style + " Range --  " + tbl.Range.ToString());
        //                   // Log.trace("STYLE: " + msg1.Name + "- " + msg1.Format.Style + " Range --  " + tbl.Rows[1].Range.ToString());

        //                    wrkRng.End = tbl.Range.End;
        //                    wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                    wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
        //                    wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                }//End if armcount
        //            }//Foreach SOA
        //                    #endregion
        //        }

        //        //Method below compares all the task in selected SOA. IF task name is Equal to CTM name or CTMName+Dose+doseunit. Get the first task event, get its visit and return visits name.

        //        private  string GetFirstTaskVisit(SOA soa, string ctmname, string ctmdose,string cmtdoseunit)
        //        { 
        //            foreach (Task tsk in soa.getTaskEnumerator().getList())
        //            {
        //                if (tsk.getActualDisplayValue().ToLower().Replace(" ", "") == (ctmname + ctmdose + cmtdoseunit).ToLower() || tsk.getActualDisplayValue().ToLower().Replace(" ", "") == (ctmname + ctmdose).ToLower() || tsk.getActualDisplayValue().ToLower().Replace(" ", "") == ctmname.ToLower())
        //                {
        //                  foreach(TaskVisit tv in soa.getTaskVisitsForTask(tsk).getList())
        //                    {                      
        //                        if (tv != null)
        //                        {  
        //                            ProtocolEvent pe = soa.getProtocolEventByID(tv.getAssociatedVisitID());
        //                            return pe.getActualDisplayValue();    //Return after getting first TaskVisit->Visit.
        //                        }
        //                    }
        //                }
        //            }
        //            return ""; //Return Empty if no events found
        //        }

        //        #endregion

        //        # region DispensingTable
        //        public void InsertDispensingTable(Word.Range wrkRng, MacrosConfig mc,string tblCaption, string filteredpurposetype,bool showtablecaption)
        //        {
        //            //This Method prints all the visit info for task-visit with purposetype = "dispensing".

        //            MacrosConfig.message msg1 = null;
        //            Word.Row newRow = null;

        //            bom_ = currdoc_.getBom();
        //            _icdSchemaMgr = currdoc_.getIcdSchemaMgr();

        //            IList ctmList = bom_.getCTMaterialEnumerator().getList();
        //            ArrayList arrTaskID = new ArrayList();
        //            ArrayList arrVisitID = new ArrayList();
        //           Hashtable htlookuptaskvisit = new Hashtable();
        //           long seltvID = 0;


        //            msg1 = mc.getMessageByName("colheader");
        //            MacrosConfig.message msg2 = mc.getMessageByName("datarow");

        //            foreach (SOA _currentSOA in bom_.getAllSchedules().getList())
        //            {

        //                //Clear the Task & Visit ArrayList, as new table for each SOA.
        //                arrTaskID.Clear();
        //                arrVisitID.Clear();


        //                foreach (TaskVisit tv in _currentSOA.getAllTaskVisits().getList())
        //                {
        //                    IEnumerator tvpenum = _currentSOA.getTaskVisitPurposes(tv);

        //                    //Loop thru each purposes of a TaskVisit, and check if one of its purpose is "Drug Dispensing"
        //                    while (tvpenum.MoveNext())
        //                    {
        //                        TaskVisitPurpose tvp = (TaskVisitPurpose)tvpenum.Current;
        //                        #region CodeforGetting_UserLabel_4_Purpose

        //                        //string tvpPurpose = currdoc_.getIcpSchemaManager().getUserLabel("PurposeTypes", tvp.getPurposeName(), tvp.getOtherPurposeText());
        //                        //if (tvpPurpose == null)
        //                        //{
        //                        //    // got to also look in our new 'dgpurposetypes' list, it is used for the 'other' purposes on right hand side on purpose tab
        //                        //    tvpPurpose = currdoc_.getIcpSchemaManager().getUserLabel(TaskVisitPurpose.DG_PREFIX + "PurposeTypes", tvp.getPurposeName(), tvp.getOtherPurposeText());
        //                        //    Log.trace("PurposeName- " + tvpPurpose);
        //                        //}
        //                        //if (LittleUtilities.isEmpty(tvp.pathToAssociatedOutcome()))
        //                        //{
        //                        //    Log.trace(tvp.getPurposeName());
        //                        //}
        //                        #endregion

        //                        if (tvp.getPurposeName() != null && tvp.getPurposeName().ToLower() == filteredpurposetype.ToLower())
        //                        {
        //                            if (arrTaskID.IndexOf(tv.getAssociatedTaskID()) < 0)
        //                            {
        //                                //Unique TASK Only
        //                                arrTaskID.Add(tv.getAssociatedTaskID());
        //                            }

        //                            if (arrVisitID.IndexOf(tv.getAssociatedVisitID()) < 0)
        //                            {
        //                                //Unique Visits Only
        //                                arrVisitID.Add(tv.getAssociatedVisitID());
        //                            }

        //                            //Add TaskVisit id as Key and "visitID+ taskID" as Value; Please verify if KEY exists has if  a TaskVisit has More then 1 purpose
        //                            //it will appear multiple time in the list.
        //                            if (!htlookuptaskvisit.ContainsKey(tv.getAssociatedVisitID() + ">" + tv.getAssociatedTaskID()))
        //                            {
        //                                //Log.trace(tv.getActualDisplayValue() + ">" + tv.getAssociatedTaskID());
        //                                htlookuptaskvisit.Add(tv.getAssociatedVisitID() + ">" + tv.getAssociatedTaskID(), tv.getObjID());
        //                            }
        //                            break;
        //                        }
        //                    }   
        //                }// End TASK VISIT


        //                if (arrTaskID.Count <= 0)
        //                {
        //                    //Print exception
        //                    wrkRng.InsertAfter("No dispensing task-visits have been defined.");
        //                    wrkRng.InsertParagraphAfter();
        //                    wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                    continue;
        //                }

        //                //Print the table


        //                wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
        //                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);


        //                #region SettingTableCaption
        //                string captiontext = tblCaption;
        //                //Custom caption if more the one SOA
        //                if (bom_.getAllSchedules().getList().Count > 1)
        //                {
        //                    if (filteredpurposetype.ToLower() == "additionalstudydrug")
        //                    {
        //                        captiontext = captiontext + " (" + _currentSOA.getActualDisplayValue() + " )"; 
        //                    }
        //                    else
        //                    {
        //                        captiontext = captiontext +" (" + _currentSOA.getActualDisplayValue() + " )";
        //                    }
        //                }


        //                #endregion

        //                Word.Table tbl = createTable(wrkRng, 1, arrTaskID.Count+1, captiontext,mc,showtablecaption);
        //                int cnt =1;
        //                int charCnt = 0;            
        //                TaskVisit seltaskvisit=null;

        //                #region HeaderRow

        //                tbl.Rows[1].Cells[1].Range.Text = "Visit ID";

        //                ArrayList arrcustomfootnotes = new ArrayList();
        //                Task tsk=null;
        //                foreach (long tskID in arrTaskID)
        //                {
        //                    cnt++;
        //                    tsk = _currentSOA.getTaskByID(tskID);
        //                    if (charCnt <= 26)
        //                    {
        //                        tbl.Rows[1].Cells[cnt].Range.Text = tsk.getActualDisplayValue() + Convert.ToChar(charCnt + 97);
        //                    }
        //                    else
        //                    {
        //                        //After a-z is done, start numbering from 1
        //                        ////For Ex: 27-26 =1, 28-26=2 etc -- > just to avoid one var
        //                        tbl.Rows[1].Cells[cnt].Range.Text = tsk.getActualDisplayValue() + Convert.ToString(charCnt-26);
        //                    }

        //                    //Make it like Footnote: Superscript.
        //                    Word.Range charRng = tbl.Rows[1].Cells[cnt].Range.Duplicate;
        //                    charRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                    charRng.Start = charRng.End - 2;
        //                    charRng.End = charRng.End - 1;
        //                    charRng.Font.Superscript = MacrosConfig.iTRUE;


        //                    charCnt++;  //For Print Characters(a -z), Footnotes
        //                    if (tsk.getValueForNode("Qtyofformulationpercontainer1289411452156") != null)
        //                    {
        //                        arrcustomfootnotes.Add(tsk.getValueForNode("Qtyofformulationpercontainer1289411452156"));
        //                    }
        //                    else
        //                    {
        //                        arrcustomfootnotes.Add("Missing Quantity for task: " + tsk.getActualDisplayValue());
        //                    }

        //                }

        //                tbl.Rows[1].HeadingFormat = VBAHelper.iTRUE;
        //#endregion


        //                //Filter Visits as they appear in soa.
        //                ArrayList orderedvisitlist = new ArrayList();
        //                foreach (ProtocolEvent pe in _currentSOA.getAllVisits().getList())
        //                {
        //                    if(arrVisitID.IndexOf(pe.getObjID()) >=0) //If Ordered 
        //                    {
        //                        orderedvisitlist.Add(pe.getObjID());
        //                    }

        //                }


        //                #region datarow                

        //                foreach (long visitID in orderedvisitlist)
        //                {
        //                    newRow = tbl.Rows.Add();
        //                    newRow.HeadingFormat = VBAHelper.iFALSE;
        //                    newRow.Cells[1].Range.Text = _currentSOA.getProtocolEventByID(visitID).getActualDisplayValue();

        //                    int taskcnt = 1;  //Cnt for columsn, which willbe Task Column.
        //                    foreach (long taskID in arrTaskID)
        //                    {
        //                      //_currentSOA.
        //                        taskcnt++;
        //                        if (htlookuptaskvisit.ContainsKey(visitID + ">" + taskID))
        //                        {
        //                            seltvID = (long)htlookuptaskvisit[visitID + ">" + taskID];
        //                            seltaskvisit = _currentSOA.getTaskVisitById(seltvID);
        //                            if (seltaskvisit != null)
        //                            {
        //                                if (seltaskvisit.getValueForNode("containersdispensedandtype1289411387406") != null)
        //                                {
        //                                 //   newRow.Cells[taskcnt].Range.Text = seltaskvisit.getValueForNode("containersdispensedandtype1289411387406").ToString();
        //                                    Formatter.FTToWordFormat2(newRow.Cells[taskcnt].Range, seltaskvisit.getValueForNode("containersdispensedandtype1289411387406").ToString());
        //                                }
        //                                else
        //                                {
        //                                    //If Custom Var is not filled.
        //                                    newRow.Cells[taskcnt].Range.Text = "Missing Containers dispensed and type.";
        //                                }
        //                            }

        //                        }
        //                        else
        //                        {
        //                            newRow.Cells[taskcnt].Range.Text = "N/A";
        //                        }
        //                    }

        //                }

        //                //Setting Style for DataRow
        //                mc.setStyle(msg2.Format.Style, currdoc_, tbl.Range);

        //                //Setting Column header Style
        //                mc.setStyle(msg1.Format.Style, currdoc_, tbl.Rows[1].Range);

        //                #endregion

        //                #region FooterROW

        //                if (arrcustomfootnotes.Count > 0)
        //                {
        //                    charCnt = 0;

        //                    //footnoterow
        //                    msg1 = mc.getMessageByName("footnoterow");

        //                    foreach (string strfootnote in arrcustomfootnotes)
        //                    {
        //                        newRow = tbl.Rows.Add();
        //                        newRow.Cells.Merge();
        //                        newRow.Range.Text = newRow.Range.Text.Replace("\r\a", "");

        //                        //Convert.ToChar(charCnt +65)
        //                        if (charCnt <= 26)
        //                        {
        //                            //newRow.Cells[1].Range.Text = Convert.ToChar(charCnt + 97) + ") " + strfootnote;
        //                            Formatter.FTToWordFormat2(newRow.Cells[1].Range, Convert.ToChar(charCnt + 97) + ") " + strfootnote);
        //                        }
        //                        else
        //                        {
        //                            //For Ex: 27-26 =1, 28-26=2 etc -- > just to avoid one var
        //                            //newRow.Cells[1].Range.Text = Convert.ToString(charCnt - 26) + ") " + strfootnote;
        //                            Formatter.FTToWordFormat2(newRow.Cells[1].Range,  Convert.ToString(charCnt - 26) + ") " + strfootnote);

        //                        }
        //                       charCnt++;

        //                       if (newRow.Cells[1].Range.Text.StartsWith("\r"))
        //                       {
        //                           newRow.Cells[1].Range.Text = newRow.Range.Text.TrimStart('\r');
        //                       }

        //                       if (charCnt > 1)
        //                       {
        //                           newRow.Borders[Word.WdBorderType.wdBorderTop].Visible = false;
        //                       }
        //                       newRow.Borders[Word.WdBorderType.wdBorderBottom].Visible = false;
        //                       newRow.Borders[Word.WdBorderType.wdBorderLeft].Visible = false;
        //                       newRow.Borders[Word.WdBorderType.wdBorderRight].Visible = false;


        //                       newRow.Range.Collapse(ref WordHelper.COLLAPSE_END);
        //                       //Setting Style for DataRow
        //                       mc.setStyle(msg1.Format.Style, currdoc_, newRow.Range);
        //                    }
        //                }

        //                #endregion

        //                tbl.Range.Collapse(ref WordHelper.COLLAPSE_END);
        //                wrkRng.End = tbl.Range.End;
        //                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
        //                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //            }
        //    }
        //        #endregion

        //        #region InvestigationalProdTable
        //        /* This method print the IP Table. If a customvariable (>Excipients> has value, attach a footnote symbol (*) in firstcell/first col.
        //         * Print value after the table.         
        //         */ 
        //        public void InsertInvestigationalProductTable(Word.Range wrkRng, MacrosConfig mc,string tblCaption,IcpInstanceManager icpinstmgr,bool showtablecaption)
        //        {

        //            //Checking if custom var is empty?
        //            bool isOther;
        //            string dose = "";
        //            string expFootnote = icpinstmgr.getDisplayValue("/FTICP/CustomVars/ExcipientFootnote","");

        //            bom_ = currdoc_.getBom();

        //            //Create table.
        //            MacrosConfig.message msg1 = null;
        //            Word.Row newRow = null;
        //            wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
        //            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

        //            Word.Table tbl = createTable(wrkRng, 1, 3, tblCaption,mc,showtablecaption);
        //            msg1 = mc.getMessageByName("colheader");
        //            MacrosConfig.message  msg2 = mc.getMessageByName("datarow");

        //            //Header Row
        //            if (!MacroBaseUtilities.isEmpty(expFootnote))
        //            {
        //                tbl.Rows[1].Cells[1].Range.Text = "Investigational Product *";

        //                //Make it like Footnote: Superscript.
        //                Word.Range charRng = tbl.Rows[1].Cells[1].Range.Duplicate;
        //                charRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                charRng.Start = charRng.End - 2;
        //                charRng.End = charRng.End - 1;
        //                charRng.Font.Superscript = MacrosConfig.iTRUE;

        //            }
        //            else
        //            {
        //                tbl.Rows[1].Cells[1].Range.Text = "Investigational Product";

        //            }



        //            tbl.Rows[1].Cells[2].Range.Text = "Dosage form and strength";
        //            tbl.Rows[1].Cells[3].Range.Text = "Manufacturer";


        //            //Setting Header style
        //            mc.setStyle(msg1.Format.Style, currdoc_, tbl.Rows[1].Range);
        //            tbl.Rows[1].HeadingFormat = VBAHelper.iTRUE;

        //            foreach (ClinicalTrialMaterial ctm in bom_.getCTMaterialEnumerator().getList())
        //            {
        //                if (ctm.getParentID()== 0)
        //                {
        //                    dose = "";
        //                    newRow = tbl.Rows.Add();
        //                    newRow.HeadingFormat = VBAHelper.iFALSE;
        //                   // newRow.Cells[1].Range.Text = Formatter.formatCleaner(ctm.getActualDisplayValue());
        //                    Formatter.FTToWordFormat2(newRow.Cells[1].Range, ctm.getActualDisplayValue());
        //                    string formulationStrength = "";

        //                    if (ctm.getFormulation() != null && ctm.getFormulation().Length > 0)
        //                    {

        //                        if (ctm.getFormulation().ToLower() == "other")
        //                        {
        //                            formulationStrength = Formatter.formatCleaner(ctm.getOtherFormulation());
        //                        }
        //                        else
        //                        {
        //                            formulationStrength = bom_.getIcpSchemaMgr().getUserLabel("FormulationTypes", ctm.getFormulation());
        //                        }

        //                      //  formulationStrength = bom_.getIcpSchemaMgr().getUserLabel("FormulationTypes", ctm.getFormulation(), ctm.getFormulation());
        //                        Log.trace("Formulation: " + formulationStrength + "   --->" + ctm.getFormulation() + " ==if other" + ctm.getOtherFormulation());
        //                    }
        //                    else
        //                    {
        //                        formulationStrength = "na";
        //                    }

        //                    //Per Resolve# 102969 
        //                    if (ctm.getValueForNode("Strength1295467858812") != null)
        //                    {
        //                        formulationStrength += " " + ctm.getValueForNode("Strength1295467858812").ToString();
        //                    }
        //                    else
        //                    {
        //                        formulationStrength += " na";
        //                    }


        //                    Formatter.FTToWordFormat2(newRow.Cells[2].Range, formulationStrength);
        //                    /* Commented due to resolve# 102969 

        //                     * */
        //                    ////if (ctm.getChildrenLikeParent().getList().Count > 0)
        //                    ////{
        //                    ////    foreach (ClinicalTrialMaterial ctmChild in ctm.getChildrenLikeParent().getList())
        //                    ////    {
        //                    ////        dose += ctmChild.getDose()  + " " + ctmChild.getDoseUnit() + ", "; 
        //                    ////    }
        //                    ////    if (dose.Trim().Length > 0)
        //                    ////    {
        //                    ////        dose = dose.Trim().TrimEnd(',');
        //                    ////    }
        //                    ////}
        //                    ////else
        //                    ////{
        //                    ////    dose = ctm.getDose() + " " + ctm.getDoseUnit();
        //                    ////}


        //                    ////newRow.Cells[2].Range.Text += dose;

        //                    if (ctm.getValueForNode("Manufacturer") != null)
        //                    {
        //                      Formatter.FTToWordFormat2( newRow.Cells[3].Range,  ctm.getValueForNode("Manufacturer").ToString());
        //                    }
        //                    //Setting Style for DataRow
        //                    mc.setStyle(msg2.Format.Style, currdoc_, newRow.Range);
        //                }
        //            }


        //            #region Footerrow

        //            msg1 = mc.getMessageByName("footnoterow");

        //            if (!MacroBaseUtilities.isEmpty(expFootnote))
        //            {
        //                newRow = tbl.Rows.Add();
        //                newRow.Cells.Merge();

        //                //There are seperate paragraph markers for each cell, so we need to bring it down to one. 
        //                //For now, we have 3 columns

        //                newRow.Range.Text.Replace("\r\a", "");

        //                //footnoterow
        //                //newRow.Cells[1].Range.Text = "* " + expFootnote;
        //                Formatter.FTToWordFormat2(newRow.Cells[1].Range, "* " + expFootnote);

        //                newRow.Borders[Word.WdBorderType.wdBorderBottom].Visible = false;
        //                newRow.Borders[Word.WdBorderType.wdBorderLeft].Visible = false;
        //                newRow.Borders[Word.WdBorderType.wdBorderRight].Visible = false;
        //                //Setting Style for DataRow
        //                mc.setStyle(msg1.Format.Style, currdoc_, newRow.Range);

        //            }

        //            #endregion

        //            wrkRng.End = tbl.Range.End;
        //            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

        //            tbl.Rows[1].HeadingFormat = VBAHelper.iTRUE;
        //            currdoc_.getActiveWordDocument().UndoClear();
        //            wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
        //            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //        }
        //        #endregion

        //        #region StudyDesignTable
        //        public void InsertStudyDesignTable(Word.Range wrkRng, MacrosConfig mc, string tblCaption, IcpInstanceManager icpinstmgr,bool showtablecaption)
        //        {
        //            MacrosConfig.message msg1 = null;
        //            Word.Row newRow = null;


        //            bom_ = currdoc_.getBom();
        //            _icdSchemaMgr = currdoc_.getIcdSchemaMgr();


        //            //Create table.
        //            wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
        //            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

        //            Word.Table tbl = createTable(wrkRng, 22, 3, tblCaption,mc,showtablecaption);


        //            int rowidx = 0;
        //            string storedval = "";

        //            //Row 0
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[1].Range.Text = "Category";
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Design Element";
        //            tbl.Rows[rowidx].Cells[3].Range.Text = "Value";

        //            //PS: Header Style & Heading Format(reprating headers across pages) appled at the bottom of method.

        //            //Row1
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[1].Range.Text = "Design";
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Phase";

        //            if (icpinstmgr.getDisplayValue(AdminDefines.PhaseCodedType, "") != null)
        //            {
        //                storedval = icpinstmgr.getDisplayValue(AdminDefines.PhaseCodedType, "");
        //                if (storedval.Length > 0)
        //                {
        //                    storedval = bom_.getIcpSchemaMgr().getUserLabel("CodedPhaseTypes", storedval, storedval);                
        //                }
        //            }
        //            if (storedval.Trim().Length<=0)
        //            {
        //                storedval = "na";
        //            }
        //            tbl.Rows[rowidx].Cells[3].Range.Text = storedval;
        //            storedval = "";

        //            //Row 2
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Purpose";

        //            if (icpinstmgr.getDisplayValue(DesignDefines.StudyPurposeType, "") != null)
        //            {
        //                storedval = icpinstmgr.getDisplayValue(DesignDefines.StudyPurposeType, "");
        //                if (storedval.Length > 0)
        //                {
        //                    storedval = bom_.getIcpSchemaMgr().getUserLabel("StudyPurposeTypes", storedval, storedval);
        //                }
        //            }
        //            if (storedval.Trim().Length <= 0)
        //            {
        //                storedval = "na";
        //            }

        //            tbl.Rows[rowidx].Cells[3].Range.Text = storedval;
        //            storedval = "";

        //            //Row 3
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Outcome";

        //            if (icpinstmgr.getDisplayValue(DesignDefines.OverallStudyOutcomeType, "") != null)
        //            {
        //                storedval = icpinstmgr.getDisplayValue(DesignDefines.OverallStudyOutcomeType, "");
        //                //OverallStudyOutcomeTypes
        //                if (storedval.Length > 0)
        //                {
        //                    storedval = bom_.getIcpSchemaMgr().getUserLabel("OverallStudyOutcomeTypes", storedval, storedval);
        //                }

        //            }
        //            if (storedval.Trim().Length <= 0)
        //            {
        //                storedval = "na";
        //            }

        //            tbl.Rows[rowidx].Cells[3].Range.Text = storedval;
        //            storedval = "";

        //            //Row 4
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Randomisation";

        //            if (icpinstmgr.getDisplayValue(DesignDefines.MethodOfAllocationType, "") != null)
        //            {
        //                storedval = icpinstmgr.getDisplayValue(DesignDefines.MethodOfAllocationType, "");

        //                     if (storedval.Length > 0)
        //                {
        //                    storedval = bom_.getIcpSchemaMgr().getUserLabel("AllocationTypes", storedval, storedval);
        //                }

        //            }
        //            if (storedval.Trim().Length <= 0)
        //            {
        //                storedval = "na";
        //            }

        //            tbl.Rows[rowidx].Cells[3].Range.Text = storedval;
        //            storedval = "";

        //            //Row 5
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Design";

        //            if (icpinstmgr.getDisplayValue(DesignDefines.StudyConfigurationType, "") != null)
        //            {
        //                storedval = icpinstmgr.getDisplayValue(DesignDefines.StudyConfigurationType, "");

        //             if (storedval.Length > 0)
        //                {
        //                    storedval = bom_.getIcpSchemaMgr().getUserLabel("ConfigurationTypes", storedval, storedval);
        //                }

        //            }
        //            if (storedval.Trim().Length <= 0)
        //            {
        //                storedval = "na";
        //            }

        //            tbl.Rows[rowidx].Cells[3].Range.Text = storedval;
        //            storedval = "";

        //            //Row 6;
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Control";

        //            if (icpinstmgr.getDisplayValue(DesignDefines.ControlType, "") != null)
        //            {
        //                storedval = icpinstmgr.getDisplayValue(DesignDefines.ControlType, "");
        //                if (storedval.Length > 0)
        //                {
        //                    storedval = bom_.getIcpSchemaMgr().getUserLabel("ControlTypes", storedval, storedval);
        //                }

        //            }
        //            if (storedval.Trim().Length <= 0)
        //            {
        //                storedval = "na";
        //            }

        //            tbl.Rows[rowidx].Cells[3].Range.Text = storedval;
        //            storedval = "";

        //            //Row 7;
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Stratification";

        //            if (icpinstmgr.getDisplayValue("/FTICP/StudyDesign/Design/Stratification", "") != null)
        //            {
        //                storedval = icpinstmgr.getDisplayValue("/FTICP/StudyDesign/Design/Stratification", "");
        //            }
        //            if (storedval.Trim().Length <= 0)
        //            {
        //                storedval = "na";
        //            }

        //            //tbl.Rows[rowidx].Cells[3].Range.Text = storedval;
        //            Formatter.FTToWordFormat2(tbl.Rows[rowidx].Cells[3].Range, storedval);
        //            storedval = "";

        //            //Row 8;
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Blinding";  //Masking

        //            if (icpinstmgr.getDisplayValue(DesignDefines.MaskingType, "") != null)
        //            {
        //                storedval = icpinstmgr.getDisplayValue(DesignDefines.MaskingType, "");
        //                //
        //                if (storedval.Length > 0)
        //                {
        //                    storedval = bom_.getIcpSchemaMgr().getUserLabel("BlindingTypes", storedval, storedval);
        //                }
        //            }
        //            if (storedval.Trim().Length <= 0)
        //            {
        //                storedval = "na";
        //            }

        //            tbl.Rows[rowidx].Cells[3].Range.Text = storedval;
        //            storedval = "";

        //            //Row 9;
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Masking";  //Refers to c

        //            if (icpinstmgr.getDisplayValue("/FTICP/StudyDesign/Design/CustomMasking", "") != null)
        //            {
        //                storedval = icpinstmgr.getDisplayValue("/FTICP/StudyDesign/Design/CustomMasking", "");
        //                if (storedval.Length > 0)
        //                {
        //                    storedval = bom_.getIcpSchemaMgr().getUserLabel("CustomMasking", storedval, storedval);
        //                }
        //            }
        //            if (storedval.Length <= 0)
        //            {
        //                storedval = "na";
        //            }

        //            tbl.Rows[rowidx].Cells[3].Range.Text = storedval;
        //            storedval = "";




        //            //Row 10;
        //            rowidx++;
        //            int duration = 0;
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Sequence and duration of periods";


        //            string strperName = "";
        //            foreach (SOA soa in bom_.getAllSchedules().getList())
        //            {
        //                foreach (Period per in soa.getPeriodEnumerator().getList())
        //                {
        //                    if (!per.isSubPeriod())
        //                    {
        //                        strperName += per.getActualDisplayValue();
        //                        //Custom duration on period
        //                        if (per.getValueForNode("Duration1295896660021") != null)
        //                        {
        //                            strperName += " " + per.getValueForNode("Duration1295896660021");
        //                        }
        //                        else
        //                        {
        //                            strperName += " na";
        //                        }
        //                        //Custom duration unit on period
        //                        if (per.getValueForNode("DurationUnit1295896676443") != null)
        //                        {
        //                            strperName += " " + per.getValueForNode("DurationUnit1295896676443");
        //                        }
        //                        else
        //                        {
        //                            strperName += " na";
        //                        }

        //                           strperName += "\r";
        //                    }
        //                }

        //            }//End For

        //            //If NO SOA or  NO Periods defined.
        //            if (strperName.Length < 0)
        //            {
        //                strperName = "na";
        //            }

        //         //   tbl.Rows[rowidx].Cells[3].Range.Text = strperName;
        //            Formatter.FTToWordFormat2(tbl.Rows[rowidx].Cells[3].Range, strperName);
        //            //Row 11
        //            #region TreatmentGroup
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Treatment Groups";
        //            string armStringlist = "";
        //            string armWeightlist = "";
        //            foreach (Arm arm in bom_.getArmEnumerator().getList())
        //            {
        //                armStringlist += arm.getActualDisplayValue();
        //                if (arm.getRandomizationWeightForArm() != null && arm.getRandomizationWeightForArm().Length > 0)
        //                {
        //                    armWeightlist += arm.getRandomizationWeightForArm() + " : ";
        //                }
        //                else
        //                {
        //                    armWeightlist += " na :";
        //                }
        //               armStringlist += "\r";
        //            }


        //            armWeightlist = armWeightlist.Trim().TrimEnd(':');  //removing the extra

        //            //tbl.Rows[rowidx].Cells[3].Range.Text = armStringlist;
        //            if (armStringlist.Trim().Length > 0)
        //            {
        //                Formatter.FTToWordFormat2(tbl.Rows[rowidx].Cells[3].Range, armStringlist);
        //            }
        //            else
        //            {
        //                tbl.Rows[rowidx].Cells[3].Range.Text = "na";
        //            }
        //            tbl.Rows[rowidx].Borders[Word.WdBorderType.wdBorderBottom].Visible = false; 

        //            //Row 11.1 
        //            rowidx++;
        //              tbl.Rows[rowidx].Cells[2].Range.Text = "(allocation ratio)";
        //              if (armWeightlist.Trim().Length > 0)
        //              {
        //                  tbl.Rows[rowidx].Cells[3].Range.Text = "(" + armWeightlist.Trim() + ")";
        //              }
        //              else
        //              {
        //                  tbl.Rows[rowidx].Cells[3].Range.Text = "na";
        //              }
        //            #endregion

        //            //Row 12
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[1].Range.Text = "Population";
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Indication";
        //            StringListHelper helper = icpinstmgr.getStringList(AdminDefines.CodedIndications, DocType.PROTOCOL);
        //            if (helper.getDisplayString().Length <= 0)
        //            {
        //                tbl.Rows[rowidx].Cells[3].Range.Text = "na";
        //            }
        //            else
        //            {
        //                tbl.Rows[rowidx].Cells[3].Range.Text = helper.getDisplayString();

        //            }
        //            //Row 13
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Age Range";

        //            string agerng = "";
        //            if (icpinstmgr.getDisplayValue("FTICP/CustomVars/MinAgeRange", "") != null)
        //            {
        //                agerng = icpinstmgr.getDisplayValue("FTICP/CustomVars/MinAgeRange", "");
        //            }
        //            else
        //            {
        //                agerng += "na";
        //            }

        //            //Min Age unit.
        //            if (icpinstmgr.getDisplayValue("FTICP/CustomVars/AgeUnit1295900040215", "") != null)
        //            {
        //                agerng += " " + icpinstmgr.getDisplayValue("FTICP/CustomVars/AgeUnit1295900040215", "");
        //            }
        //            else
        //            {
        //                agerng += " na";
        //            }


        //            agerng += " to ";

        //            if (icpinstmgr.getDisplayValue("FTICP/CustomVars/MaxAgeRange", "") != null)
        //            {
        //                agerng += icpinstmgr.getDisplayValue("FTICP/CustomVars/MaxAgeRange", "");
        //            }
        //            else
        //            {
        //                agerng += " na";
        //            }

        //            //Max Age unit.
        //            if (icpinstmgr.getDisplayValue("FTICP/CustomVars/AgeUnit1295900063013", "") != null)
        //            {
        //                agerng += " " + icpinstmgr.getDisplayValue("FTICP/CustomVars/AgeUnit1295900063013", "");
        //            }
        //            else
        //            {
        //                agerng += " na";
        //            }
        //            tbl.Rows[rowidx].Cells[3].Range.Text = agerng;

        //            ////CountryCV cve = Tspd.Bridge.BridgeProxy.getInstance().getCountryList();
        //            ////StringListHelper cnthelper = icpinstmgr.getStringList(AdminDefines.CodedCountriesType, DocType.PROTOCOL);

        //            ////foreach (StringListHelper.CodedValue  cv in cnthelper.getCodedValueArray())
        //            ////{
        //            ////    Log.trace(cve.findByAttributeValue("name", cv.Value)[0].ToString());
        //            ////}
        //            ////tbl.Rows[rowidx].Cells[3].Range.Text = helper.getDisplayString();


        //            //Row 14
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[1].Range.Text = "Logistics";
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Approximate Number of centres";
        //            if (icpinstmgr.getDisplayValue(AdminDefines.PlannedNumberStudyCenters, "") != null)
        //            {
        //                storedval = icpinstmgr.getDisplayValue(AdminDefines.PlannedNumberStudyCenters, "");
        //            }
        //            if (storedval.Trim().Length <= 0)
        //            {
        //                storedval = "na";
        //            }

        //            tbl.Rows[rowidx].Cells[3].Range.Text = storedval;
        //            storedval = "";

        //            //row 15
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Multi Country";

        //            StringListHelper countrylist = icpinstmgr.getStringList(AdminDefines.CodedCountriesType, DocType.PROTOCOL);
        //            if (countrylist.getCodedValueArray().Count > 1)
        //            {
        //                tbl.Rows[rowidx].Cells[3].Range.Text = "Yes";
        //            }
        //            else
        //            {
        //                tbl.Rows[rowidx].Cells[3].Range.Text = "No";
        //            }

        //            //row 16
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Approximate Number of randomised/treated subjects";
        //            tbl.Rows[rowidx].Cells[2].WordWrap = true;

        //            if (icpinstmgr.getDisplayValue("FTICP/Administrative/ProtocolSkeleton/ofrandomizedortreatedsubjects1294774451421", "") != null)
        //            {
        //                storedval = icpinstmgr.getDisplayValue("FTICP/Administrative/ProtocolSkeleton/ofrandomizedortreatedsubjects1294774451421", "");
        //            }
        //            if (storedval.Trim().Length <= 0)
        //            {
        //                storedval = "na";
        //            }

        //            //tbl.Rows[rowidx].Cells[3].Range.Text = storedval;
        //            Formatter.FTToWordFormat2(tbl.Rows[rowidx].Cells[3].Range, storedval);
        //            storedval = "";


        //            //row 17
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Approximate min number of subjects per centre";

        //            if (icpinstmgr.getDisplayValue("FTICP/Administrative/ProtocolSkeleton/Minofsubjectspercenter", "") != null)
        //            {
        //                storedval = icpinstmgr.getDisplayValue("FTICP/Administrative/ProtocolSkeleton/Minofsubjectspercenter", "");
        //            }
        //            else
        //                if (storedval.Trim().Length <= 0)
        //                {
        //                    storedval = "na";
        //                }

        //            Formatter.FTToWordFormat2(tbl.Rows[rowidx].Cells[3].Range, storedval);
        //            storedval = "";

        //            //row 18
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Approximate max number of subjects per centre";

        //            if (icpinstmgr.getDisplayValue("FTICP/Administrative/ProtocolSkeleton/Maxofsubjectspercenter", "") != null)
        //            {
        //                storedval = icpinstmgr.getDisplayValue("FTICP/Administrative/ProtocolSkeleton/Maxofsubjectspercenter", "");
        //            }

        //                if (storedval.Trim().Length <= 0)
        //                {
        //                    storedval = "na";
        //                }

        //                Formatter.FTToWordFormat2(tbl.Rows[rowidx].Cells[3].Range, storedval);
        //            storedval = "";

        //            //row 19
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[1].Range.Text = "Other";
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Data Monitoring Committee";

        //            if (icpinstmgr.getDisplayValue("FTICP/Administrative/ProtocolSkeleton/DMC1294778558687", "") != null)
        //            {
        //                storedval = icpinstmgr.getDisplayValue("FTICP/Administrative/ProtocolSkeleton/DMC1294778558687", "");
        //                if (storedval.Length > 0)
        //                {
        //                    storedval = bom_.getIcpSchemaMgr().getUserLabel("yesNo", storedval);
        //                }

        //            }
        //            if (storedval.Trim().Length <= 0)
        //            {
        //                storedval = "na";
        //            }

        //            Formatter.FTToWordFormat2(tbl.Rows[rowidx].Cells[3].Range, storedval);
        //            storedval = "";

        //            //row 20
        //            rowidx++;
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Timing of interim analyses";

        //            if (icpinstmgr.getDisplayValue("FTICP/CustomVars/InterimAnalyis", "") != null)
        //            {
        //                storedval = icpinstmgr.getDisplayValue("FTICP/CustomVars/InterimAnalyis", "");
        //            }
        //            if (storedval.Trim().Length <= 0)
        //            {
        //                storedval = "na";
        //            }

        //            Formatter.FTToWordFormat2(tbl.Rows[rowidx].Cells[3].Range, storedval);
        //            storedval = "";

        //            wrkRng.End = tbl.Range.End;
        //            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //            wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
        //            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

        //            msg1 = mc.getMessageByName("datarow");
        //            mc.setStyle(msg1.Format.Style, currdoc_, tbl.Range);


        //            tbl.Rows[1].HeadingFormat = VBAHelper.iTRUE;
        //            //Setting Header Style
        //            msg1 = mc.getMessageByName("rowheader");
        //            mc.setStyle(msg1.Format.Style, currdoc_, tbl.Rows[1].Range); 



        //        }
        //        #endregion 

        //        #region EmergencyContactsTable


        //        public void InsertEmergencyContacts(Word.Range wrkRng, MacrosConfig mc, string tblCaption,bool showtablecaption)
        //        {
        //            MacrosConfig.message msg1 = null;

        //            bom_ = currdoc_.getBom();
        //            IList cenum = bom_.getContactEnumerator().findByAttributeValue("EmergencyContact1288639952250", "Y");

        //            if (cenum.Count == 0)
        //            {
        //                Log.trace("No Emergency Contacts Found.");
        //                return;
        //            }

        //            Word.Table tbl = createTable(wrkRng, cenum.Count + 1, 3, tblCaption,mc,showtablecaption);
        //            int rowidx = 1;
        //            string roletype="";
        //            string address = "";

        //            string filepath = currdoc_.getTrialProject().getTemplateDirPath() + "\\dyntmplts\\mailingFormat.xml";

        //            Log.trace(filepath);

        //            # region HeaderRow and Caption
        //            tbl.Rows[rowidx].Cells[1].Range.Text = "Name";
        //            tbl.Rows[rowidx].Cells[2].Range.Text = "Role in the study";
        //            tbl.Rows[rowidx].Cells[3].Range.Text = "Address & telephone number";

        //            tbl.Rows[rowidx].HeadingFormat = VBAHelper.iTRUE;

        //            rowidx++;
        //           //Please see below for Header Style. 
        //            #endregion




        //            AddressFormat aformat = new AddressFormat(filepath);
        //            foreach (Contact contact in cenum)
        //            {

        //                Formatter.FTToWordFormat2(tbl.Rows[rowidx].Cells[1].Range, contact.getPersonName());
        //                roletype = bom_.getIcpSchemaMgr().getUserLabel("ContactRoleTypes", contact.getRoleType());
        //                if (roletype.Length > 0)
        //                {
        //                    Formatter.FTToWordFormat2(tbl.Rows[rowidx].Cells[2].Range, roletype);
        //                }

        //              IList addobj=  aformat.format("", "", contact.getAddress1(), contact.getAddress2(), contact.getCity(), contact.getStateProv(), 
        //                    contact.getPostalCode(),contact.getCountry(), contact.getTel(), contact.getFax(), contact.getEmail(),"Multiline");

        //              address="";
        //              foreach (AddressFormat.AddressLine l in addobj)
        //              {
        //                  address += l.text +"\r";
        //              }

        //              address = address.TrimEnd('\r');
        //              tbl.Rows[rowidx].HeadingFormat = VBAHelper.iFALSE;
        //                Formatter.FTToWordFormat2(tbl.Rows[rowidx].Cells[3].Range, address);
        //                rowidx++;
        //            }

        //            wrkRng.End = tbl.Range.End;
        //            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //            wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
        //            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

        //            msg1 = mc.getMessageByName("datarow");
        //            mc.setStyle(msg1.Format.Style, currdoc_, tbl.Range);

        //            //Applying Header style to Row 1
        //            msg1 = mc.getMessageByName("rowheader");
        //            mc.setStyle(msg1.Format.Style, currdoc_, tbl.Rows[1].Range); 
        //        }

        //        #endregion

        //        #region Print_Sponsor_Org

        //        public void InsertSponsorOrg(Word.Range wrkRng, MacrosConfig mc)
        //        {
        //            MacrosConfig.message msg1 = null;

        //            bom_ = currdoc_.getBom();

        //            IList org_enum = bom_.getOrganizationEnumerator().findByAttributeValue("organizationType", "sponsor");
        //            if (org_enum.Count == 0)
        //            {
        //                Log.trace("No sponsor organization found.");
        //                return;
        //            }

        //            string filepath = currdoc_.getTrialProject().getTemplateDirPath() + "\\dyntmplts\\mailingFormat.xml";
        //            Log.trace(filepath);
        //            AddressFormat aformat = new AddressFormat(filepath);
        //            msg1 = mc.getMessageByName("normalstyle");
        //            string address = "";
        //            foreach (Organization org in org_enum)
        //            {
        //                IList addobj = aformat.format("", org.getName(), org.getAddress1(), org.getAddress2(), org.getCity(), org.getStateProv(), org.getPostalCode(), org.getCountry(), org.getTel(), org.getFax(), org.getEmail(),"SingleLine");
        //                address = "";
        //                foreach (AddressFormat.AddressLine l in addobj)
        //                {
        //                    if (l.text.Trim() != ",")
        //                    {
        //                        address += l.text;
        //                    }
        //                }

        //                Formatter.FTToWordFormat2(wrkRng, address);
        //                wrkRng.InsertParagraphAfter();
        //                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                mc.setStyle(msg1.Format.Style, currdoc_, wrkRng); 
        //            }

        //        }
        //        #endregion

        //        #region Print_Lead_Investigator
        //        public void InsertLeadInvestigator(Word.Range wrkRng, MacrosConfig mc, IcpInstanceManager icpinstmgr)
        //        {
        //            MacrosConfig.message msg1 = null;
        //            bom_ = currdoc_.getBom();

        //            //Assumption, templates are set so that there will always be one selection.
        //            string inv_type = icpinstmgr.getDisplayValue("/FTICP/StudyConduct/SCG/LeadInvestigator", "");
        //            string inv_userlabel = bom_.getIcpSchemaMgr().getUserLabel("LeadInvestigator", inv_type);


        //            //Printing Heading
        //            try
        //            {
        //                //code for setting style.
        //                string headerstyle = "A-Unassigned";
        //                mc.setStyle(headerstyle, currdoc_, wrkRng);
        //            }
        //            catch (Exception e)
        //            { }
        //            wrkRng.InsertAfter(inv_userlabel);
        //            wrkRng.InsertParagraphAfter();
        //            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);


        //             string filepath = currdoc_.getTrialProject().getTemplateDirPath() + "\\dyntmplts\\mailingFormat.xml";
        //            Log.trace(filepath);
        //            AddressFormat aformat = new AddressFormat(filepath);

        //            //Resetting it back to Normal.
        //            msg1 = mc.getMessageByName("normalstyle");
        //            mc.setStyle(msg1.Format.Style, currdoc_, wrkRng);


        //            //Loop thru each contact and check if it has matching user label as the inv_userlabel.
        //            //Print all info for that contact. and insert Paragraph for next one

        //            bool hasContact = false;
        //            IList cenum = bom_.getContactEnumerator().getList();

        //            int contactcnt = 0;
        //            foreach (Contact c in cenum)
        //            {
        //                if (c.getRoleType() != null && c.getRoleType().Length>0)
        //                {
        //                    string curruserlabel = bom_.getIcpSchemaMgr().getUserLabel("ContactRoleTypes", c.getRoleType(), c.getRoleType());
        //                    if (curruserlabel.ToLower().Equals(inv_userlabel.ToLower()))
        //                    {
        //                        hasContact = true;
        //                        string orgname="";
        //                        foreach (Organization o in bom_.getAssociatedOrganizations(c).getList())
        //                        {
        //                            orgname = o.getActualDisplayValue();   //Print the first organization.
        //                            break;
        //                        }

        //                        //bom_.getOrganizationEnumerator



        //                        IList addobj = aformat.format(c.getActualDisplayValue(), orgname, c.getAddress1(), c.getAddress2(), c.getCity(), c.getStateProv(),
        //                   c.getPostalCode(), c.getCountry(), c.getTel(), c.getFax(), c.getEmail(),"Multiline");


        //                        foreach (AddressFormat.AddressLine l in addobj)
        //                        {
        //                          //  wrkRng.InsertAfter(l.text + "\v");
        //                            Formatter.FTToWordFormat2(wrkRng, l.text + "\v");
        //                            //wrkRng.InsertParagraphAfter();

        //                            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                            mc.setStyle(l.style, currdoc_, wrkRng);
        //                        }

        //                        if (cenum.Count != contactcnt)
        //                        {
        //                            wrkRng.InsertParagraphAfter();
        //                            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                        }
        //                    }
        //                }
        //            }

        //            if (!hasContact)
        //            {
        //                wrkRng.InsertAfter("No " + inv_userlabel + " has been specified.");
        //                wrkRng.InsertParagraphAfter();
        //                wrkRng.Collapse(ref WordHelper.COLLAPSE_END); 
        //            }


        //        }
        //        #endregion

        //        #region InsertAnalysisVarbyType
        //        public void  InsertAnalysisVariableByType(Word.Range wrkRng, MacrosConfig mc, IcpInstanceManager icpinstmgr,string outcometype)
        //        {
        //            MacrosConfig.message msg1 = null;
        //            bom_ = currdoc_.getBom();
        //            String varName="",varDesc="";
        //            IList outcomelist = bom_.getOutcomes().findByAttributeValue("outcomeType", outcometype);
        //            if (outcomelist.Count < 0)
        //            {
        //                //Print Message "No Outcomes with selected Outcome type" 
        //                return;
        //            }
        //            ArrayList varinOutcome = new ArrayList();

        //            MacrosConfig.message  msg = mc.getMessageByName("liststyle");

        //             foreach(Outcome o in outcomelist)
        //             {
        //               MappedVariableCollection mvc = o.getVariableCollection();
        //               IEnumerator refs = mvc.getVarRefs();
        //               while (refs.MoveNext())
        //               {
        //                   VarRef aRef = (VarRef)refs.Current;
        //                   long varID = aRef.getVariableID();
        //                   varName = "";
        //                   varDesc = "";

        //                   if (!varinOutcome.Contains(varID))
        //                   {
        //                   mc.setStyle(msg.Format.Style, currdoc_, wrkRng);
        //                   varinOutcome.Add(varID);  //Adding VarId to Local list for this macro.So if a var is associate with 2 outcomes of same type its not repeated
        //                   StudyVariable var = bom_.getVariableDictionary().findByICPID(varID);
        //                   varName = var.getActualDisplayValue();
        //                   varDesc = var.getFullDescription();

        //                   varName = varName.Replace("\n", "\v");
        //                   varDesc = varDesc.Replace("\n", "\v");

        //                   //combining into one string.
        //                   if (varDesc != null && varDesc.Length > 0)
        //                   {
        //                       varName = varName + " - " + varDesc;
        //                   }
        //                   Formatter.FTToWordFormat2(wrkRng, varName);
        //                   //wrkRng.InsertAfter(var.getActualDisplayValue() + " - " + var.getFullDescription());
        //                   wrkRng.InsertParagraphAfter();
        //                   wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                   Log.trace("Outcome: " + o.getActualDisplayValue() + "Variable - " + var.getActualDisplayValue());
        //                   }
        //               }
        //             }





        //        }
        //        #endregion

        //        #region Print_Stats_Model

        //        public void InsertStatsModel(Word.Range wrkRng,MacrosConfig mc)
        //        {
        //             MacrosConfig.message msg1 = null;
        //            bom_ = currdoc_.getBom();
        //            string modelfulldesc="";
        //            ///Get all Types based on Enum type : AnalysisTypes. List it in the same order as they appear in Stats Model.

        //              ArrayList modeltypes = bom_.getIcpSchemaMgr().getEnumPairs("AnalysisTypes");
        //              foreach (EnumPair ep in modeltypes)
        //              {
        //                  if (ep.getSystemName().ToLower() == "other")
        //                  {
        //                      continue;
        //                  }


        //                  IList statenum = bom_.getStatsMgr().getAllStatisticalModelEnumerator().findByAttributeValue("analysisType", ep.getSystemName());
        //                  if (statenum.Count > 0)
        //                  {
        //                      msg1 = mc.getMessageByName("statisticsheaderstyle");
        //                      //Only Print header if there are any models
        //                      mc.setStyle(msg1.Format.Style, currdoc_, wrkRng);                      
        //                      wrkRng.InsertAfter(Formatter.stripFormatInstruction(ep.getUserLabel()));
        //                      wrkRng.InsertParagraphAfter();
        //                      wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

        //                      msg1 = mc.getMessageByName("statsbodytextstyle");

        //                      //Print the list
        //                      foreach (StatisticalModel smodel in statenum)
        //                      {
        //                          mc.setStyle(msg1.Format.Style, currdoc_, wrkRng);
        //                          modelfulldesc = smodel.getFullDescription();
        //                          Formatter.FTToWordFormat2(wrkRng, modelfulldesc);
        //                          wrkRng.InsertParagraphAfter();
        //                          wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

        //                      }
        //                  }

        //              }
        //                //Print "OTHER" - If any.

        //            ///Step 1: Get the list of all other stats models and UNIQUE getOtherModel
        //                ArrayList otherlist = new ArrayList();
        //                IList statmodelList = bom_.getStatsMgr().getAllStatisticalModelEnumerator().findByAttributeValue("analysisType/@codeValue", "other");
        //                foreach (StatisticalModel smodel in statmodelList)
        //                {
        //                   // smodel.getOtherModelValue()
        //                    if(!otherlist.Contains(smodel.getOtherModelValue()))
        //                    {
        //                        otherlist.Add(smodel.getOtherModelValue());
        //                    }                  
        //                }

        //                string headerText = "";
        //                foreach (string strOthermodel in otherlist)
        //                {
        //                    IList otherstatsenum = bom_.getStatsMgr().getAllStatisticalModelEnumerator().findByAttributeValue("analysisType", strOthermodel);

        //                    if (otherstatsenum.Count > 0)
        //                    {

        //                        //Print the Header
        //                        mc.setStyle("Heading 3", currdoc_, wrkRng);

        //                        //Resolve: 107774  -- Convert first char to Uppercase
        //                        headerText = "";
        //                        headerText = Formatter.stripFormatInstruction(strOthermodel).ToLower();
        //                        headerText = headerText.Substring(0, 1).ToUpper() + headerText.Substring(1);
        //                        wrkRng.InsertAfter(headerText);
        //                        wrkRng.InsertParagraphAfter();
        //                        wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

        //                        //Print the list
        //                        foreach (StatisticalModel smodel in otherstatsenum)
        //                        {
        //                            mc.setStyle("Normal", currdoc_, wrkRng);
        //                            modelfulldesc = smodel.getFullDescription();
        //                            Formatter.FTToWordFormat2(wrkRng, modelfulldesc);
        //                            wrkRng.InsertParagraphAfter();
        //                            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                        }
        //                    }
        //                }
        //        }
        //        #endregion

        //        #region Print_Population_sets

        //        public void InsertPopulationSets(Word.Range wrkRng, MacrosConfig mc)
        //        {
        //            MacrosConfig.message msg1 = null;
        //            bom_ = currdoc_.getBom();
        //            string fulldesc = "";
        //            ///Get all Types based on Enum type : AnalysisSetTypes. List it in the same order as they appear in Stats Model.

        //            ArrayList modeltypes = bom_.getIcpSchemaMgr().getEnumPairs("AnalysisSetTypes");
        //            foreach (EnumPair ep in modeltypes)
        //            {
        //                if (ep.getSystemName().ToLower() == "other")
        //                {
        //                    continue;
        //                }

        //                IList popstat_enum = bom_.getStatsMgr().getAnalysisPopulationSetEnumerator().findByAttributeValue("setType", ep.getSystemName());

        //                if (popstat_enum.Count> 0)               
        //                {
        //                    msg1 = mc.getMessageByName("statisticsheaderstyle");
        //                    //Only Print header if there are any models
        //                    mc.setStyle(msg1.Format.Style, currdoc_, wrkRng);
        //                    wrkRng.InsertAfter(ep.getUserLabel());
        //                    wrkRng.InsertParagraphAfter();
        //                    wrkRng.Collapse(ref WordHelper.COLLAPSE_END);


        //                    msg1 = mc.getMessageByName("statsbodytextstyle");
        //                    //Print the list
        //                    foreach (AnalysisPopulationSet obj_pop_set in popstat_enum)
        //                    {
        //                        mc.setStyle(msg1.Format.Style, currdoc_, wrkRng);
        //                        fulldesc = obj_pop_set.getFullDescription();
        //                        Formatter.FTToWordFormat2(wrkRng, fulldesc);
        //                        wrkRng.InsertParagraphAfter();
        //                        wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

        //                    }
        //                }
        //            }


        //            //Print "OTHER" - If any.

        //            ///Step 1: Get the list of all other analysis Sets and UNIQUE getOtherPopset
        //            ArrayList otherlist = new ArrayList();
        //            IList pop_set_List = bom_.getStatsMgr().getAnalysisPopulationSetEnumerator().findByAttributeValue("setType/@codeValue", "other");
        //            foreach (AnalysisPopulationSet pop_set in pop_set_List)
        //            {
        //                // smodel.getOtherModelValue()
        //                if (!otherlist.Contains(pop_set.getOtherAnalysisSetValue()))
        //                {
        //                    otherlist.Add(pop_set.getOtherAnalysisSetValue());
        //                }
        //            }

        //            string headerText = "";
        //            foreach (string strOther in otherlist)
        //            {
        //                IList otherpop_set_enum = bom_.getStatsMgr().getAnalysisPopulationSetEnumerator().findByAttributeValue("setType", strOther);

        //                if (otherpop_set_enum.Count > 0)
        //                {

        //                    //Print the Header
        //                    mc.setStyle("Heading 3", currdoc_, wrkRng);
        //                    //Resolve: 107774 - Convert first char it to upper case
        //                    headerText = "";
        //                    headerText = Formatter.stripFormatInstruction(strOther).ToLower();
        //                    headerText = headerText.Substring(0, 1).ToUpper() + headerText.Substring(1);
        //                    wrkRng.InsertAfter(headerText);

        //                    wrkRng.InsertParagraphAfter();
        //                    wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

        //                    //Print the list
        //                    foreach (AnalysisPopulationSet popset in otherpop_set_enum)
        //                    {
        //                        mc.setStyle("Normal", currdoc_, wrkRng);
        //                        fulldesc = popset.getFullDescription();
        //                        Formatter.FTToWordFormat2(wrkRng, fulldesc);
        //                        wrkRng.InsertParagraphAfter();
        //                        wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

        //                    }
        //                }
        //            } //End For each - Other
        //        }
        //        #endregion

        //        #region Print_Analyses

        //        public void InsertAnalyses(Word.Range wrkRng, MacrosConfig mc)
        //        {
        //            MacrosConfig.message msg1 = null;
        //            bom_ = currdoc_.getBom();
        //            string fulldesc = "";
        //            ///Get all Types based on Enum type : AnalysisSetTypes. List it in the same order as they appear in Stats Model.

        //            ArrayList modeltypes = bom_.getIcpSchemaMgr().getEnumPairs("AnalysisRoles");
        //            foreach (EnumPair ep in modeltypes)
        //            {
        //                if (ep.getSystemName().ToLower() == "other")
        //                {
        //                    continue;
        //                }


        //                IList analysis_enum = bom_.getStatsMgr().getAnalysisEnumerator().findByAttributeValue("analysisRole", ep.getSystemName());

        //                if (analysis_enum.Count > 0)
        //                {
        //                    msg1 = mc.getMessageByName("statisticsheaderstyle");
        //                    //Only Print header if there are any models
        //                    mc.setStyle(msg1.Format.Style, currdoc_, wrkRng);
        //                    wrkRng.InsertAfter(ep.getUserLabel());
        //                    wrkRng.InsertParagraphAfter();
        //                    wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

        //                    msg1 = mc.getMessageByName("statsbodytextstyle");
        //                    //Print the list
        //                    foreach (Analysis ana in analysis_enum)
        //                    {
        //                        mc.setStyle(msg1.Format.Style, currdoc_, wrkRng);
        //                        fulldesc = ana.getFullDescription();
        //                        if (fulldesc.Length <= 0)
        //                        {
        //                            fulldesc = "Instructions missing for Analysis: " + ana.getBriefDescription();
        //                        }
        //                        Formatter.FTToWordFormat2(wrkRng, fulldesc);
        //                        wrkRng.InsertParagraphAfter();
        //                        wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

        //                    }
        //                }
        //            }


        //            //Print "OTHER" - If any.

        //            ///Step 1: Get the list of all other analysis Sets and UNIQUE getOtherPopset
        //            ArrayList otherlist = new ArrayList();
        //            IList analysis_List = bom_.getStatsMgr().getAnalysisEnumerator().findByAttributeValue("analysisRole/@codeValue", "other");
        //            foreach (Analysis ana in analysis_List)
        //            {
        //                // smodel.getOtherModelValue()
        //                if (!otherlist.Contains(ana.getOtherAnalysisValue()))
        //                {
        //                    otherlist.Add(ana.getOtherAnalysisValue());
        //                }
        //            }

        //            string headerText = "";
        //            foreach (string strOther in otherlist)
        //            {
        //                IList otheranalysis_enum = bom_.getStatsMgr().getAnalysisEnumerator().findByAttributeValue("analysisRole", strOther);

        //                if (otheranalysis_enum.Count > 0)
        //                {

        //                    //Print the Header
        //                    mc.setStyle("Heading 3", currdoc_, wrkRng);
        //                    //Resolve: 107774 - Convert first char it to upper case
        //                    headerText = "";
        //                    headerText = Formatter.stripFormatInstruction(strOther).ToLower();
        //                    headerText = headerText.Substring(0, 1).ToUpper() + headerText.Substring(1);
        //                    wrkRng.InsertAfter(headerText);

        //                    wrkRng.InsertParagraphAfter();
        //                    wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

        //                    //Print the list
        //                    foreach (Analysis ana in otheranalysis_enum)
        //                    {
        //                        mc.setStyle("Normal", currdoc_, wrkRng);
        //                        fulldesc = ana.getFullDescription();
        //                        Formatter.FTToWordFormat2(wrkRng, fulldesc);
        //                        wrkRng.InsertParagraphAfter();
        //                        wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                    }
        //                }
        //            } //End For each - Other
        //        }
        //        #endregion

        //        #region commoncode
        //        public  Word.Table createTable(Word.Range viewRng, int rows, int cols,string tblcaption,MacrosConfig mc,bool showtablecaption)
        //        {

        //            // Turn off auto caption for Word tables.
        //            Word.AutoCaption ac = wdApp_.AutoCaptions.get_Item(ref WordHelper.AUTO_CAPTION_WORD_TABLE);
        //            bool oldState = ac.AutoInsert;
        //            ac.AutoInsert = false;

        //            Word.Range wrkRng = viewRng.Duplicate;

        //            //IF showtablecaption is set then only show caption, else skip it. If "table caption text" is not there, then insert "Table 1"

        //            if (showtablecaption)
        //            {

        //                //Applying Caption Style
        //                MacrosConfig.message m1 = mc.getMessageByName("tablecaptionstyle");

        //                if (tblcaption != null)
        //                {

        //                    if (tblcaption.Length > 0)
        //                    {
        //                        if (m1.Text.ToLower() == "space")
        //                        {
        //                            tblcaption = " " + tblcaption;
        //                        }
        //                        else if (m1.Text.ToLower() == "tab")
        //                        {
        //                            tblcaption = "\t" + tblcaption;
        //                        }
        //                    }

        //                    object tableCaption = tblcaption;

        //                    wrkRng.InsertAfter(" "); // single space which'll get shifted after the caption.
        //                    Word.Range captionRng = wrkRng.Duplicate;
        //                    captionRng.Collapse(ref WordHelper.COLLAPSE_START);

        //                    // put the caption in if the TableView gives you one...
        //                    if (tableCaption != null)
        //                    {
        //                        captionRng.InsertCaption(
        //                            ref WordHelper.CAPTION_LABEL_TABLE, ref tableCaption,
        //                            ref VBAHelper.OPT_MISSING, ref WordHelper.CAPTION_POSITION_ABOVE, ref VBAHelper.OPT_MISSING);
        //                    }

        //                    // Mark the added single space and replace it with a paragraph mark.
        //                    wrkRng.Start = wrkRng.End - 1;
        //                    wrkRng.InsertParagraph();

        //                    //Setting configured style for Caption, so List of Tables can be configured.
        //                    mc.setStyle(m1.Format.Style, currdoc_, wrkRng);

        //                    // Collapse the range to the end of the paragraph mark so that the table can be added
        //                    // after it.
        //                    wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
        //                }

        //                wdDoc.UndoClear();

        //            }//End if showcaption


        //            // Collapse the range to the end of the paragraph mark so that the table can be added
        //            // after it.
        //            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);


        //            // Insert the table. Note, that the table is inserted starting at but after the range.
        //            // So viewRng isn't increased.
        //            Word.Table tbl = wdDoc.Tables.Add(
        //                wrkRng, rows, cols,
        //                ref WordHelper.WORD8_TABLE_BEHAVIOR, ref VBAHelper.OPT_MISSING);


        //            tbl.Borders.Enable = VBAHelper.iTRUE;

        //            // Reinstate auto caption for Word tables.
        //            ac.AutoInsert = oldState;

        //           // tbl.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthAuto;
        //            //tbl.PreferredWidth = (float)100.00;

        //            tbl.AllowAutoFit = true;
        //            tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);
        //            ////Added per Matt's request to allow the column's Automatically set its width.
        //          //  tbl.Columns.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthAuto;



        //            //tbl.Columns[1].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
        //            //tbl.Columns[1].PreferredWidth = tbl.Application.InchesToPoints(0.84f);  //Fixed width

        //            // Increase viewRng to include the table.
        //            viewRng.End = tbl.Range.End;

        //            viewRng.Collapse(ref WordHelper.COLLAPSE_END);

        //            return tbl;
        //        }
        //# endregion


    }
}

