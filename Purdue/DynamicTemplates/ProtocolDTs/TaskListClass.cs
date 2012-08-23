
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Tspd.Utilities;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Tspddoc;
using Tspd.MacroBase;
using Tspd.Macros;
using System.Windows.Forms;


namespace TspdCfg.Purdue.DynTmplts
{
    public class TaskListClass
    {
        public BusinessObjectMgr _bom = null;
        public IcdSchemaManager _icdSchemaMgr = null;


        #region data
        public class TaskList
        {
            public string Name { get; set; }
            public string Type { get; set; }
            public string Lab { get; set; }
            public ArrayList  value;
            public ArrayList ListofTask;
        }
        
        public List<TaskList> TaskObjects = new List<TaskList>();


        public void AddItem(string Name, string ListType, string Lab, ArrayList value, ArrayList ListofTask)
        {
            try
            {
                TaskList objTlist = new TaskList();
                objTlist.Name = Name;
                objTlist.Type = ListType;
                objTlist.Lab = Lab;
                objTlist.value = value;
                objTlist.ListofTask = ListofTask;
                TaskObjects.Add(objTlist);                
            }
            catch (Exception e)
            {
                
            }
        }

        public void FillTaskList()
        {
            try
            {
                IList soaList = _bom.getAllSchedules().getList();
                Period parentPerorSubper = null;
             
                    foreach (SOA soa in soaList)
                    {
                        Log.trace("Schedule:  " + soa.getActualDisplayValue());
                        IList tskList = soa.getTaskEnumerator().getList();
                        foreach (Task tsk in tskList)
                        {
                            TaskVisitEnumerator en = new TaskVisitEnumerator(soa.getTaskVisitForTaskID(tsk.getObjID()),_icdSchemaMgr.getTVTemplate());

                            foreach(TaskVisit tv in en.getList())
                            {                           
                                EventScheduleBase visit = soa.getVisitOfTaskVisit(tv);
                                //Check if sub period exists, the return its ScheduleItemtype:
                                 parentPerorSubper =  soa.getParentOfScheduleItem(visit);
                                 if (parentPerorSubper.isSubPeriod())
                                 {
                                     //Checking epoch for SubPeriod
                                     UpdateTaskVisitList4Epoch(tsk, GetEpochofPeriodSubPeriod(parentPerorSubper),tv.isCentralFacility(),tv.isLocalFacility());
                                     Log.trace("SubPeriod:: " + parentPerorSubper.getActualDisplayValue());
                                     parentPerorSubper = soa.getParentOfScheduleItem(parentPerorSubper);  //Get Period;
                                 }

                                 UpdateTaskVisitList4Epoch(tsk, GetEpochofPeriodSubPeriod(parentPerorSubper), tv.isCentralFacility(), tv.isLocalFacility());
                                 Log.trace("Period:: " + parentPerorSubper.getActualDisplayValue());

                                 if (tv.hasPurposes())
                                {
                                    IEnumerator tvPurpose = soa.getTaskVisitPurposes(tv);
                                    while (tvPurpose.MoveNext())
                                    {
                                        TaskVisitPurpose tvp = (TaskVisitPurpose)tvPurpose.Current;
                                        if (!LittleUtilities.isEmpty(tvp.pathToAssociatedOutcome()))
                                        {
                                            Outcome o1 = _bom.getOutcomeByPath(tvp.pathToAssociatedOutcome());
                                            if (o1 != null)
                                            {
                                                if (UpdateTaskVisitList4purpose(tsk, o1.getOutcomeType(),tv.isCentralFacility(),tv.isLocalFacility()))
                                                {
                                                    break;  //once task is no need to loop for next purpose.
                                                }
                                            }
                                        }
                                    }//End While
                                 }                               
                            } //End While 
                        } //End Task
                    }//End SOA                

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private string GetEpochofPeriodSubPeriod(Period _selPeriod)
        {
            //Returns the Schedule Item Type AKA Epoch
            return _selPeriod.getScheduleItemType();
   
        }

        private bool UpdateTaskVisitList4Epoch(Task tsk,string visitEpochType,bool CentralLab,bool LocalLab)
        {
            bool _taskAdded = false;
            var tasklist = from task in TaskObjects
                           //where task.Name..Equals("Pinal") & task.Type[1].Equals("Alpha")
                           where (task.Type.Equals("epoch") && task.value.IndexOf(visitEpochType)>=0)
                           select new { objTask = task };


            foreach (var item in tasklist)
            {
                if (!item.objTask.ListofTask.Contains(tsk))
                {
                    if (item.objTask.Lab == "exclude")
                    {  //Central & Local Labs should be "false"
                        if (CentralLab == false && LocalLab == false)
                        {
                            item.objTask.ListofTask.Add(tsk);
                            _taskAdded = true;
                        }
                    }
                    else if (item.objTask.Lab == "only")
                    {
                        if (CentralLab == true || LocalLab == true)
                        { //Local Labs only should be true.
                            item.objTask.ListofTask.Add(tsk);
                            _taskAdded = true;
                        }
                    }
                    else if (item.objTask.Lab == "include")
                    {   
                            item.objTask.ListofTask.Add(tsk);
                            _taskAdded = true;
                    }
                }
            }

            return _taskAdded;
        }

        private bool UpdateTaskVisitList4purpose(Task tsk, string purposeType,bool CentralLab,bool LocalLab)
        {
            bool _taskAdded = false;
            var tasklist = from task in TaskObjects
                           //where task.Name..Equals("Pinal") & task.Type[1].Equals("Alpha")
                           where (task.Type.Equals("purpose") && task.value.IndexOf(purposeType) >= 0)
                           select new { objTask = task };
            

            foreach (var item in tasklist)
            {
                if (!item.objTask.ListofTask.Contains(tsk))
                {
                    if (item.objTask.Lab == "exclude")
                    {  //Central & Local Labs should be "false"
                        if (CentralLab == false && LocalLab == false)
                        {
                            item.objTask.ListofTask.Add(tsk);
                            _taskAdded = true;
                        }
                    }
                    else if (item.objTask.Lab == "only")
                    {
                        if (CentralLab == true || LocalLab == true)
                        { //Local Labs only should be true.
                            item.objTask.ListofTask.Add(tsk);
                            _taskAdded = true;
                        }
                    }
                    else if (item.objTask.Lab == "include")
                    {
                        item.objTask.ListofTask.Add(tsk);
                        _taskAdded = true;
                    }
                }
            }

            return _taskAdded;
        }



        #region Task_based_on_Epoch

    
        public void TaskwithTaskEvents()
        {
            try
            {
                ArrayList taskwithPE = new ArrayList();
                SOAEnumerator soaEnum = _bom.getAllSchedules();
                while (soaEnum.MoveNext())
                {
                   SOA soa = soaEnum.getCurrent();
                   IList taskList = soa.getTaskEnumerator().getList();

                   foreach (Task tsk in taskList)
                   {
                       if (soa.getAllVisitsOfTask(tsk).Count > 0)  //Filter task having Task Events
                       {
                           if(!taskwithPE.Contains(tsk.getObjID()))
                           {
                               taskwithPE.Add(tsk.getObjID());  //Add Task if not in list.
                           }
                       }
                   }//End For
                }//End While
            }
            catch (Exception e)
            { 
            } 
        }

        private ArrayList Get_VisitwithTaskEvents(SOA _currentSOA)
        {
            PeriodEnumerator perEnum = _currentSOA.getPeriodEnumerator();
            IList VisitswithEvents = new ArrayList();
            IList visitList = new ArrayList();
            Hashtable sortedPerEnum = new Hashtable();
            ArrayList ai = new ArrayList();
            while (perEnum.MoveNext())
            {
                Period per = (Period)perEnum.Current;
                if (_currentSOA.getSubPeriodCount(per) == 0)
                {
                    //EventScheduleEnumerator VisitEnum = _currentSOA.getPeriodChildren(per);
                    visitList = _currentSOA.getPeriodChildren(per).getList();
                    foreach (EventScheduleBase visit in visitList)
                    {
                        if (VisitswithEvents.IndexOf(visit.getObjID()) >= 0)
                        {
                            if (taskVisitExists(_currentSOA,visit))
                            {
                                VisitswithEvents.Add(visit.getObjID());
                            }
                        } 
                    }
                }//End IF

                else
                {   //Getting sub period.
                    EventScheduleEnumerator subPerChildren = _currentSOA.getPeriodChildren(per);
                    while (subPerChildren.MoveNext())
                    {
                        visitList = _currentSOA.getPeriodChildren(per).getList();
                        foreach (EventScheduleBase visit in visitList)
                        {
                            if (VisitswithEvents.IndexOf(visit.getObjID()) >= 0)
                            {
                                if (taskVisitExists(_currentSOA, visit))
                                {
                                    VisitswithEvents.Add(visit.getObjID());
                                }
                            } 
                        }
                    }
                }//end else

                visitList.Clear(); //Clear the visit list.
              
            }//end while

            return ai;
        }

        private bool taskVisitExists(SOA _soa,EventScheduleBase _visit)
        {  //This methods, gets an Visits and returns if there are any Task Events.
            TaskVisitEnumerator en = new TaskVisitEnumerator(_soa.getTaskVisitForVisitID(_visit.getObjID()), _icdSchemaMgr.getTVTemplate());
            if (en.getList().Count > 0)
            {
                return true;  //If there are Task Visit for selected Visit
            }
            return false;
        }

        #endregion

        #endregion

    }
}
