#define xuseMonths

using System;
using System.Collections;

using Tspd.Tspddoc;
using Tspd.MacroBase;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;

using System.Diagnostics;

namespace TspdCfg.Pfizer.DynTmplts
{
	public class PfizerUtil
	{
		public static readonly string NORMAL = "Normal";

		public static readonly string PFIZER_STYLE_TEXT_NUM = "Text:Num";
		public static readonly string PFIZER_STYLE_TABLETEXT_10 = "TableText:10";
		public static readonly string PFIZER_STYLE_TABLETEXT_12 = "TableText:12";
		public static readonly string PFIZER_STYLE_TEXT_TI12 = "Text:Ti12";		
		public static readonly string PFIZER_STYLE_TEXT_TI12_LEFT = "Text:Ti12 + Left";
		public static readonly string PFIZER_STYLE_TEXT_BULL = "Text:Bull";
		public static readonly string PFIZER_STYLE_TABLETEXT_BULL_10 = "TableText:Bull:10";

		public static readonly string PFIZER_INC_CRIT_NUMBERED_LISTB = "TSPDIncCritListB";
		public static readonly string PFIZER_INC_CRIT_NUMBERED_LISTS = "TSPDIncCritListS";

		public static readonly string PFIZER_EXCL_CRIT_NUMBERED_LISTB = "TSPDExclCritListB";
		public static readonly string PFIZER_EXCL_CRIT_NUMBERED_LISTS = "TSPDExclCritListS";

		public static readonly string INCLUSION = "Inclusion";
		public static readonly string EXCLUSION = "Exclusion";

		public class TimeUnit
		{
			private long _multiplier;
			private string _systemName;

			static ArrayList _timeUnits;

			public static readonly long MINUTES = 1;
			public static readonly string sMINUTES = "minutes";

			public static readonly long HOURS = MINUTES * 60;
			public static readonly string sHOURS = "hours";

			public static readonly long DAYS = HOURS * 24;
			public static readonly string sDAYS = "days";

			public static readonly long WEEKS = DAYS * 7;
			public static readonly string sWEEKS = "weeks";

#if useMonths
			public static readonly long MONTHS = DAYS * 30;
			public static readonly string sMONTHS = "months";
#endif

			static TimeUnit() 
			{
				_timeUnits = new ArrayList();

				_timeUnits.Add(new TimeUnit(MINUTES, sMINUTES));
				_timeUnits.Add(new TimeUnit(HOURS, sHOURS));
				_timeUnits.Add(new TimeUnit(DAYS, sDAYS));
				_timeUnits.Add(new TimeUnit(WEEKS, sWEEKS));
#if useMonths
				_timeUnits.Add(new TimeUnit(MONTHS, sMONTHS));
#endif
			}

			public static TimeUnit find(string sTimeUnit) 
			{
				if (MacroBaseUtilities.isEmpty(sTimeUnit))
				{
					return null;
				}

				foreach (TimeUnit tu in _timeUnits)
				{
					if (tu._systemName.Equals(sTimeUnit)) 
					{
						return tu;
					}
				}

				return null;
			}

			public static TimeUnit getMoreGranular(TimeUnit tu) 
			{
				int i = _timeUnits.IndexOf(tu);
				if (i == 0 || i == -1) 
				{
					return null;
				}

				return _timeUnits[i - 1] as TimeUnit;
			}

			public static TimeUnit getMin() 
			{
				return _timeUnits[0]  as TimeUnit;
			}

			public static TimeUnit getMax() 
			{
				return _timeUnits[_timeUnits.Count - 1]  as TimeUnit;
			}

			TimeUnit(long multiplier, string systemName) 
			{
				_multiplier = multiplier;
				_systemName = systemName;
			}

			public long getMultiplier() 
			{
				return _multiplier;
			}

			public string getSystemName() 
			{
				return _systemName;
			}
		}


		public class PeriodAndVisit
		{
			public Period per = null;
			public ProtocolEvent visit = null;
			public ArrayList tcList = new ArrayList();
		}

		public class PeriodAndVisitComparer : IComparer
		{
			#region IComparer Members

			int IComparer.Compare(Object x, Object y)  
			{
				PeriodAndVisit pv1 = x as PeriodAndVisit;
				PeriodAndVisit pv2 = y as PeriodAndVisit;
				
				int perCompare = pv1.per.getSequence().CompareTo(pv2.per.getSequence());

				if (perCompare != 0) return perCompare;

				int visitCompare = pv1.visit.getSequence().CompareTo(pv2.visit.getSequence());

				return visitCompare;
			}

			#endregion
		}	

		public PfizerUtil() {}

		public static ArrayList getPeriodVisitList(SOA soa, long armID, ArrayList orderedTopLevelEvents) 
		{
			ArrayList pvList = new ArrayList();

			foreach (EventScheduleBase obj in orderedTopLevelEvents)
			{
				Period per = obj as Period;
				if (per == null) 
				{
					continue;
				}
				
				ArrayList visits = PfizerUtil.getVisits(soa, armID, per, EventType.EventSubType.Scheduled);

				foreach (ProtocolEvent visit in visits) 
				{
					PfizerUtil.PeriodAndVisit pv = new PfizerUtil.PeriodAndVisit();
					pv.per = per;
					pv.visit = visit;

					pvList.Add(pv);
				}
			}

			return pvList;
		}

		public static void addTimeUnit(ref long timeInMinutes, 
			string sTime, string sTimeUnit, 
			out bool isBadTime, out bool isBadTimeUnit) 
		{
			isBadTimeUnit = false;
			long duration = PfizerUtil.getNumber(sTime, out isBadTime);

			TimeUnit tu = TimeUnit.find(sTimeUnit);
			if (tu == null)
			{
				isBadTimeUnit = true;
				return;
			}

			timeInMinutes += duration * tu.getMultiplier();
		}

		public static string getDisplayTime(long timeInMinutes, string sTimeUnit) 
		{
			long remainder = timeInMinutes;
			string s = "";

			TimeUnit minTu = TimeUnit.getMin();
			TimeUnit tu = TimeUnit.find(sTimeUnit);
			if (tu == null) 
			{
				return "Invalid TimeUnit: " + sTimeUnit;
			}

			while (remainder >= minTu.getMultiplier() && tu != null) 
			{
				string  dispTime = getDisplayTime(remainder, out remainder, tu);	
				if (dispTime.Length > 0)
				{
					s += " " + dispTime;
				}
				tu = TimeUnit.getMoreGranular(tu);
			}

			return s;
		}

		private static string getDisplayTime(long timeInMinutes, out long remainder, TimeUnit tu) 
		{
			string displayTimeUnit = tu.getSystemName();
			long displayTime = Math.DivRem(timeInMinutes, tu.getMultiplier(), out remainder);

			if (displayTime == 0) 
			{
				return ""; // PA reinstated this return type so we could skip "leading 0 units"
			}

			string disp = displayTime.ToString();
			
			// trim off s
			if (displayTime == 1) 
			{
				displayTimeUnit = displayTimeUnit.Substring(0, displayTimeUnit.Length - 1); 
			}


			return disp + " " + displayTimeUnit;
		}

		public static long getNumber(string s, out bool isBad) 
		{
			long n = 0;
			isBad = false;

			try 
			{
				n = long.Parse(s);
			}
			catch (Exception ex) 
			{
				isBad = true;
			}

			return n;
		}

		/// <summary>
		/// returns an ordered array of the ProtocolEvents that have duration
		/// in the given arm.  This method works for schema and non-schema trials.
		/// </summary>
		/// <param name="arm"></param>
		/// <returns></returns>
		public static ArrayList getAllPlannedVisitsByArm(SOA soa, long arm)
		{
			ArrayList retArray = new ArrayList();
			ProtocolEventEnumerator en = null;
			try
			{
				ArrayList orderedTopLevelEvents = new ArrayList();
				soa.getTopLevelActivityList(arm, null, orderedTopLevelEvents);
				foreach (EventScheduleBase obj in orderedTopLevelEvents)
				{
					if (obj is Period) 
					{
						EventScheduleEnumerator evts = soa.getPeriodChildren(obj as Period);
						while(evts.MoveNext())
						{
							EventScheduleBase evb = evts.getCurrent();
							if (evb is ProtocolEvent)
							{
								if (((ProtocolEvent)evb).getEventType().getSubtype() ==
									EventType.EventSubType.Scheduled) 
								{
									retArray.Add(evb);
								}
							}
							else
							{
								// its a sub period
								EventScheduleEnumerator subEvts = soa.getPeriodChildren(evb as Period);
								while(subEvts.MoveNext())
								{
									EventScheduleBase evb2 = subEvts.getCurrent();
									if (((ProtocolEvent)evb2).getEventType().getSubtype() ==
										EventType.EventSubType.Scheduled) 
									{
										retArray.Add(evb2);
									}
								}
							}
						}
					}
				}
			}
			catch(Exception ex)
			{
				Log.exception(ex, "Creating Ordered Visit List for Arm");
			}

			return retArray;
		}

		public static string getCellText(Word.Cell c) 
		{
			Word.Range cellRange = c.Range.Duplicate;
			cellRange.End--;

			string cellText = cellRange.Text;
			if (cellText == null)
			{
				cellText = "";
			}

			return cellText;
		}

		public static ArrayList getVisits(SOA soa, long armID, Period per) 
		{
			ArrayList visits = new ArrayList();

			EventScheduleEnumerator subPerChildren = soa.getPeriodChildren(per);
			while (subPerChildren.MoveNext())
			{
				EventScheduleBase visit = subPerChildren.getCurrent();
				visits.Add(visit);
			}

			return visits;
		}

		public static  ArrayList getVisits(SOA soa, long armID, Period per, EventType.EventSubType evtSubType) 
		{
			ArrayList visits = new ArrayList();
			ArrayList allVisits = getVisits(soa, armID, per);
			
			foreach (ProtocolEvent visit in allVisits) 
			{
				if (visit != null && visit.getEventType().getSubtype() == evtSubType) 
				{
					visits.Add(visit);
				}
			}

			return visits;
		}

		public class TCWrapper 
		{
			public TimeChunk tc;

			public string Label = "";
			public string Start = "";
			public string End = "";
			public string Unit = "";
			// this is needed when a time chunk is given a time or span that matches another day
			// some schedules might specify 8h on day 2 of the period, while another schedule might say 32h
			public int	  DayIndex = 0;

			public TCWrapper(TimeChunk tc, string Start, string End, string Unit) 
			{
				this.tc = tc;
				this.Label = this.Start = Start;
				this.End = End;
				this.Unit = Unit;

				if (this.End.Length != 0) 
				{
					this.Label = this.Start + " - " + this.End;
				}
			}

			public bool isSpan() 
			{
				return this.End.Length != 0;
			}

			public long getStartMinute() 
			{
				long start = -1;
				try
				{
					double start1 = double.Parse(Start);
					TimeUnit tu1 = TimeUnit.find(Unit);
					start1 *= tu1.getMultiplier();
					start  = (long)start1;
				}
				catch (Exception)
				{
					start = -1;
				}

				return start;
			}

			public long getEndMinute() 
			{
				long end = -1;
				try
				{
					double end1 = double.Parse(End);
					TimeUnit tu1 = TimeUnit.find(Unit);
					end1 *= tu1.getMultiplier();
					end = (long)end1;
				}
				catch(Exception)
				{
					end = -1;
				}
				return end;
			}

			/// <summary>
			/// tells us that the TaskVisit has a span of > 24 hours.  This may mean
			/// that in an SOA that regards visit=day by duration, then there will probably
			/// be several task visits with this same span label  e.g.  48-96.  These will
			/// end up being merged into a single ribbon.
			/// </summary>
			/// <returns></returns>
			public bool isMultiDaySpan()
			{
				return ((getEndMinute() - getStartMinute())) > TimeUnit.DAYS;
			}
		}

		public class TCWrapperComparer : IComparer
		{
			#region IComparer Members

			int IComparer.Compare(Object x, Object y)  
			{
				TCWrapper tcw1 = x as TCWrapper;
				TCWrapper tcw2 = y as TCWrapper;

				long start1 = tcw1.getStartMinute();
				long start2 = tcw2.getStartMinute();

				// in case user repeat timechunk labels on different days.
				int dayindex1 = tcw1.DayIndex;
				int dayindex2 = tcw2.DayIndex;
				int comp = dayindex1.CompareTo(dayindex2);
				if (comp == 0) 
				{		
					return start1.CompareTo(start2);
				}
				else
				{
					return comp;
				}
			}

			#endregion
		}

		public class MergePair
		{
			public string txt = "";
			public int r;
			public int c1;
			public int c2;

			public MergePair(string txt, int r, int c1, int c2) 
			{
				this.txt = txt;
				this.r = r;
				this.c1 = c1;
				this.c2 = c2;
			}

			public static void merge(Word.Table tbl, ArrayList merge) 
			{
				ArrayList mergeList = new ArrayList(merge);
				mergeList.Reverse();

				// clear the list
				merge.Clear();

				Word.Row row;
				Word.Cell c1;
				Word.Cell c2;

				foreach (PfizerUtil.MergePair mp in mergeList) 
				{
					row = tbl.Rows.Item(mp.r);

					c1 = row.Cells.Item(mp.c1);
					c2 = row.Cells.Item(mp.c2);

					c1.Merge(c2);
					c1.Range.Text = mp.txt;
				}
			}
		}

		public static bool parseTimePoint(string sTp, out string startTime, out string endTime, out string unit, out string serr) 
		{
			sTp = sTp.Trim();

			startTime = "";
			endTime = "";
			bool haveSpan = false;
			unit = TimeUnit.sMINUTES;
			serr = "";

			int lastPos = 0;

			try 
			{
				bool curNegative = false;
				string curNumber = "";
				string curUnit = "";

				for (int i = 0; i < sTp.Length; i++) 
				{
					lastPos = i;
					char c = sTp[i];

					if (char.IsDigit(c) || c == '.') 
					{
						curNumber += c;
						continue;
					}

					if (char.IsLetter(c)) 
					{
						curUnit += c;

						if (curNumber.Length == 0) 
						{
							continue;
						}

						string saveUnit = curUnit;
						if (saveNumber(ref curNegative, ref curNumber, ref curUnit, ref startTime, ref endTime, ref haveSpan, ref serr)) 
						{
							curUnit = saveUnit;
							continue;
						}

						return false;
					}

					if (c == '-') 
					{
						// Negative
						if (curNumber.Length == 0) 
						{
							if (i < sTp.Length-1 && char.IsDigit(sTp[i + 1])) 
							{
								curNegative = true;
								continue;
							}
						}
					}

					// - 'to'
					if (c == '-' || curUnit == "to") 
					{
						if (haveSpan) 
						{
							serr = "already have span (- or to)";
							return false;
						}

						haveSpan = true;

						if (curNumber.Length == 0) 
						{
							curNegative = false;
							curNumber = "";
							curUnit = "";
							continue;
						}

						if (saveNumber(ref curNegative, ref curNumber, ref curUnit, ref startTime, ref endTime, ref haveSpan, ref serr))
						{
							continue;
						}

						return false;
					}

					if (char.IsWhiteSpace(c))
					{
						if (curNumber.Length == 0) 
						{
							continue;
						}

						if (saveNumber(ref curNegative, ref curNumber, ref curUnit, ref startTime, ref endTime, ref haveSpan, ref serr))
						{
							continue;
						}

						return false;
					}
				}

				if (curUnit.Length != 0)
				{
					if (curUnit == "m" || curUnit == "min" || curUnit == "minutes") 
					{
						unit = TimeUnit.sMINUTES;
					}
					else if (curUnit == "h" || curUnit == "hour" || curUnit == "hours")
					{
						unit = TimeUnit.sHOURS;
					}
					else
					{

						serr = "invalid time unit: " + curUnit;
						return false;
					}
				}

				// Left over number
				if (curNumber.Length != 0) 
				{
					if (!saveNumber(ref curNegative, ref curNumber, ref curUnit, ref startTime, ref endTime, ref haveSpan, ref serr))
					{
						return false;
					}
				}

				if (startTime.Length == 0) 
				{
					serr = "no start time specified";
					return false;
				}

				if (haveSpan && endTime.Length == 0) 
				{
					serr = "no start time specified";
					return false;
				}
			}
			catch (Exception ex) 
			{
				serr = ex.Message;
				return false;
			}

			return true;
		}

		private static bool saveNumber(ref bool curNegative, ref string curNumber, ref string curUnit, ref string startTime, ref string endTime, ref bool haveSpan, ref string serr)
		{
			double d = double.Parse(curNumber);
			if (curNegative) d = -d;

			if (startTime.Length == 0) 
			{
				startTime = d.ToString();

				curNegative = false;
				curNumber = "";
				curUnit = "";
				return true;
			}

			if (endTime.Length == 0) 
			{
				if (!haveSpan) 
				{
					if (!curNegative) 
					{
						serr = "missing span (- or to)";
						return false;
					}

					haveSpan = true;
					d = -d;
				}

				endTime = d.ToString();

				curNegative = false;
				curNumber = "";
				curUnit = "";

				double startT = double.Parse(startTime);

				if (startT > d) 
				{
					serr = "start > end time";
					return false;
				}

				if (startT == d) 
				{
					serr = "start == end time";
					return false;
				}

				return true;
			}

			serr = "start and end time are already defined";
			return false;
		}

		// This is a copy of PEExploder.getTimeChunks,
		// It was copied here due to bugs in core code
		public static ArrayList getTimeChunks(SOA soa, ProtocolEvent pe)
		{
			TaskVisitEnumerator te = soa.getTaskVisitsForVisit(pe);
			Hashtable buckets = new Hashtable(10);
			ArrayList bucketOrder = new ArrayList(20);
			ArrayList tempList = new ArrayList();

			foreach (TaskVisit tv in te.getList())
			{
				try
				{
					tempList.Add(tv);
				}
				catch(Exception ex)
				{
					Log.exception(ex, "Duplicate sequence in Task Sequencing");
				}

				TaskVisitEnumerator children = soa.getCopiesOfTaskVisit(tv);
				foreach (TaskVisit tvSub in children.getList())
				{
					try
					{
						tempList.Add(tvSub);
					}
					catch(Exception ex)
					{
						Log.exception(ex, "Duplicate sequence in Task Sequencing");
					}
				}
			}

			tempList.Sort(new BusinessObjectFactory.SequenceSort());
			TimeChunk currentBucket = null;
			foreach (TaskVisit tv2 in tempList)
			{
				TimeChunk tc = null;
				if (tv2.getLabel() != null && tv2.getLabel().Length > 0)
				{
					if (buckets[tv2.getLabel()] == null)
					{
						tc = new TimeChunk(tv2.getLabel());
						buckets[tv2.getLabel()] = currentBucket = tc;
						bucketOrder.Add(tc);
					}
					else 
					{
						currentBucket = (TimeChunk)buckets[tv2.getLabel()];
					}
				}
				else if (currentBucket == null)
				{
					currentBucket = new TimeChunk("0h");
					currentBucket.setDefaultLabel(true);

					bucketOrder.Add(currentBucket);
				}
				currentBucket.addTaskVisit(tv2);
			}

			if (Log.shouldLog(TraceLevel.Info))
			{
				foreach(TimeChunk tc in bucketOrder)
				{
					Log.trace("Got bucket: " + tc.Label);
					foreach (TaskVisit ttv in tc.TaskVisitsInOrder)
					{
						Log.trace("TaskVisit: " + ttv.getObjID());
					}
				}
			}
			return bucketOrder;
		}
		//kludge for the regimen dosing task template in 285 [May 2007]: avoid if possible....
		public static Task getTaskByName(SOA soa, String name)
		{
			TaskEnumerator taskEnum = soa.getTaskEnumerator();
			while (taskEnum.MoveNext()) 
			{
				Task task = taskEnum.getCurrent();
				if(task.getBriefDescription().CompareTo(name) == 0)
					return task;
			}
			return null;
		}

		private static ClinicalTrialMaterial findCTM(BusinessObjectMgr bom, long ctmID) 
		{
			CTMaterialEnumerator ctEnum = bom.getCTMaterialEnumerator();
			while (ctEnum.MoveNext()) 
			{
				ClinicalTrialMaterial ctm = ctEnum.getCurrent();
				if (ctm.getObjID() == ctmID) 
				{
					return ctm;
				}
			}

			return null;
		}

		public static IList findCTMsByType(BusinessObjectMgr bom, string role) 
		{
			ArrayList ctmList = new ArrayList();
			CTMaterialEnumerator ctEnum = bom.getCTMaterialEnumerator();
			while (ctEnum.MoveNext()) 
			{				
				ClinicalTrialMaterial ctm = ctEnum.getCurrent();
				string primaryRole = ctm.getPrimaryRole();
					
				if (!MacroBaseUtilities.isEmpty(primaryRole) && primaryRole.Equals(role))
				{
					ctmList.Add(ctm);
				}
			}
			return ctmList;
		}


		//
		public static IList getDosingTaskskByTAType(SOA soa, String role, BusinessObjectMgr bom, IcdSchemaManager icdSchemaMgr)
		{
			IList taskList = soa.getTaskEnumerator().getList();
			IList dosingTasks = new ArrayList();
			foreach (Task task in taskList) 
			{
				if(task.isDosingTask() == true)
				{
					DosingTask dosingTask = new DosingTask(
						task.getObjectRoot(), icdSchemaMgr.getTemplateByClass(typeof(DosingTask)));
					long ctmID = dosingTask.getctMaterialID();
					ClinicalTrialMaterial ctm = findCTM(bom, ctmID);
					string ctmRole = ctm.getPrimaryRole();
					if (!MacroBaseUtilities.isEmpty(ctmRole) && ctmRole.Equals(role) == true) 
					{
						dosingTasks.Add(dosingTask);
					}
				}
			}
			return dosingTasks;
		}
	}
}
