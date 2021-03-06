#define xuseMonths

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using Tspd.Tspddoc;
using Tspd.MacroBase;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;

using System.Diagnostics;

using Word = Microsoft.Office.Interop.Word;

namespace TspdCfg.Purdue.DynTmplts
{
	public class PurdueUtil
	{
		public static readonly string NORMAL = "Normal";


        public static readonly string PURDUE_STYLE_TABLETEXT_10 = "Table Text";
        public static readonly string PURDUE_STYLE_FOOTNOTE = "Table Footnote";


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

			public static readonly long SECONDS = 1;
			public static readonly string sSECONDS = "seconds";

			public static readonly long MINUTES = SECONDS * 60;
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

				_timeUnits.Add(new TimeUnit(SECONDS, sSECONDS));
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

		public PurdueUtil() {}

      

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
				
				ArrayList visits = PurdueUtil.getVisits(soa, armID, per, EventType.EventSubType.Scheduled);

				foreach (ProtocolEvent visit in visits) 
				{
					PurdueUtil.PeriodAndVisit pv = new PurdueUtil.PeriodAndVisit();
					pv.per = per;
					pv.visit = visit;

					pvList.Add(pv);
				}
			}

			return pvList;
		}

		public static void addTimeUnit(ref long timeInSeconds, string sTime, out bool isBadTime) 
		{
			long duration = PurdueUtil.getNumber(sTime, out isBadTime);

			if (isBadTime) 
			{
				return;
			}

			TimeUnit tu = TimeUnit.find(TimeUnit.sSECONDS);
			if (tu == null)
			{
				return;
			}

			timeInSeconds += duration * tu.getMultiplier();
		}


        public static Word.Style getStyle(Word.Document thisDoc_, string styleName)
        {
            Word.Style wordStyle = null;
            try
            {
                object oStyName = (object)styleName;
                wordStyle = thisDoc_.Styles[oStyName];
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.ToString());
            }
            return wordStyle;

        }


        public static void resartListNumber(Word.Range wrkRng)
        {
           
                object applyTo = Word.WdListApplyTo.wdListApplyToWholeList;
                object behaviour = Word.WdDefaultListBehavior.wdWord9ListBehavior;
                int npara = wrkRng.ListParagraphs.Count;


                if (wrkRng.ListFormat.ListType == Word.WdListType.wdListOutlineNumbering)
                {
                      if(wrkRng.ListFormat.ListTemplate.ListLevels.Count>0)
                        wrkRng.ListFormat.ApplyListTemplate(wrkRng.ListFormat.ListTemplate, ref VBAHelper.oFALSE, ref applyTo, ref WordHelper.WORD9_LIST_BEHAVIOR);
                }
                else
                {
                    try
                    {

                        if (npara > 0)
                        {  //If you remove this loop, it will err if no paragraph is there.

                            wrkRng.ListParagraphs[1].Range.ListFormat.ApplyListTemplate(
                         wrkRng.ListParagraphs[1].Range.ListFormat.ListTemplate,
                         ref VBAHelper.oFALSE,
                         ref applyTo, ref WordHelper.WORD9_LIST_BEHAVIOR);

                        }
                    }
                    catch (Exception e1)
                    
                    {
                        Log.exception(e1, e1.StackTrace + e1.Message);
                    }

                }//endelse

        }
		public static string getDisplayTime(long timeInSeconds, string sTimeUnit) 
		{
			long remainder = timeInSeconds;
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

		private static string getDisplayTime(long timeInSeconds, out long remainder, TimeUnit tu) 
		{
			string displayTimeUnit = tu.getSystemName();
			long displayTime = Math.DivRem(timeInSeconds, tu.getMultiplier(), out remainder);

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

				foreach (PurdueUtil.MergePair mp in mergeList) 
				{
					row = tbl.Rows[mp.r];

					c1 = row.Cells[mp.c1];
					c2 = row.Cells[mp.c2];

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

        private static Treatment findTreatment(BusinessObjectMgr bom, long treatmentID) 
		{
            return bom.getTreatment(treatmentID);
		}

        public class TreatmentComponentAndTestArticle
        {
            public Treatment MatchingTreatment;
            public Component MatchingComponent;
            public TestArticle MatchingTestArticle;
        }

        public static List<TreatmentComponentAndTestArticle> findTreatmentsByRole(BusinessObjectMgr bom, string role) 
		{
            List<TreatmentComponentAndTestArticle> list = new List<TreatmentComponentAndTestArticle>();
            IEnumerable<BaseProtocolObject> treatments = bom.getTreatments().Enumerable;
            foreach (Treatment treatment in treatments.OfType<Treatment>().OrderBy(tr => tr.getSequence()))
            {
                foreach (Component component in bom.getAssociatedComponents(treatment).Enumerable)
                {
                    TestArticle testArticle = bom.getTestArticle(component.AssociatedTestArticleID);
                    if ((testArticle != null) && role.Equals(testArticle.PrimaryRole))
                        list.Add(
                            new TreatmentComponentAndTestArticle
                            {
                                MatchingTreatment = treatment,
                                MatchingComponent = component,
                                MatchingTestArticle = testArticle
                            });
                }
            }
            return list;
        }

		//
		public static IList getDosingTaskskByTAType(SOA soa, String role, BusinessObjectMgr bom, IcdSchemaManager icdSchemaMgr)
		{
			IList taskList = soa.getTaskEnumerator().getList();
			IList dosingTasks = new ArrayList();
            /* There are no dosing tasks in TSPD 3.1 - LAP
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
            */
			return dosingTasks;
		}

        public static int GetRange(Word.Range _selRng, string findString)
        {
            try
            {
                bool found = false;
                object _optMissing = System.Reflection.Missing.Value;

                Word.Range sel_ = _selRng;
                sel_.Find.Text = findString.ToString();
                sel_.Find.Forward = true;
                sel_.Find.Wrap = Word.WdFindWrap.wdFindContinue;
                sel_.Find.MatchCase = false;
                sel_.Find.MatchWholeWord = false;
                sel_.Find.MatchWildcards = false;

                found = sel_.Find.Execute(
                      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing,
                      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing,
                      ref _optMissing, ref _optMissing, ref _optMissing,
                      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing);

                if (found)
                {
                    if (findString.Length == 4)
                    {
                        return sel_.Start;
                    }
                    else
                    {
                        return sel_.End;
                    }

                }
                else
                {
                    return -1;
                }

            }
            catch (Exception ex)
            {

            }

            return -1;
        }

        private static Word.Range ApplyandStripFormatting(Word.Range modRng)
        {
            //this Method will first apply FT formatting, and then remove the Tags. Some of the values are Hard coded
            //due to nature.
            string displayText = modRng.Text.Trim();
            string startTag = displayText.Substring(0, 6);
            string endTag = displayText.Substring(displayText.Length - 7, 7);
            string tag = "";
            if (startTag.Equals(endTag.Replace("/", "")))
            {
                tag = startTag.Substring(4, 1);
                switch (tag)
                {
                    case "i":
                        modRng.Font.Italic = VBAHelper.iTRUE;
                        break;
                    case "u":
                        modRng.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                        break;
                    case "b":
                        modRng.Font.Bold = VBAHelper.iTRUE;
                        break;
                    case "-":
                        modRng.Font.Subscript = VBAHelper.iTRUE;
                        break;
                    case "+":
                        modRng.Font.Superscript = VBAHelper.iTRUE;
                        break;
                    default:
                        break;
                }

                //removing formatting tags.
                Word.Range tmpRng = modRng.Duplicate;
                tmpRng.Collapse(ref WordHelper.COLLAPSE_START);
                tmpRng.End = tmpRng.End + 6;
                tmpRng.Text = "";

                tmpRng = modRng.Duplicate;
                tmpRng.Collapse(ref WordHelper.COLLAPSE_END);
                tmpRng.Start = tmpRng.Start - 7;
                tmpRng.Text = "";

            }



            return modRng;
        }

        public static void setStyle(TspdDocument thisDoc_, string styleName, Word.Range _selRng)
        {
            try
            {
                thisDoc_.getStyleHelper().setNamedStyle(styleName, _selRng);
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.ToString());
            }

        }

        public static void RemoveFTOrphanTags(Word.Range selRng, String findString)
        {
            try
            {
                bool found = false;
                object _optMissing = System.Reflection.Missing.Value;
                object replace = Word.WdReplace.wdReplaceAll;
                Word.Range sel_ = selRng;
                sel_.Find.Text = findString.ToString();
                sel_.Find.Forward = true;
                sel_.Find.Replacement.Text = "";
                sel_.Find.Wrap = Word.WdFindWrap.wdFindContinue;
                    sel_.Find.MatchCase = false;
                sel_.Find.MatchWholeWord = false;
                sel_.Find.MatchWildcards = false;

                found = sel_.Find.Execute(
                      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing,
                      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing,
                      ref _optMissing, ref _optMissing, ref replace,
                      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing);


            }
            catch (Exception ex)
            {
            }

        }

        public static string StripFTFormats(string str)
        {
            str = str.Replace("<ft:b>", "");
            str = str.Replace("<ft:i>", "");
            str = str.Replace("<ft:u>", "");
            str = str.Replace("<ft:->", "");
            str = str.Replace("<ft:+>", "");
            return str;
        }

        public static void FormatString(TspdDocument tspdDoc_, Word.Range wrkRng)
        {
            //   string currText ="";
            int newEndRng = wrkRng.End;
            string str = "", cleanStr = "";
            Word.Range storedRange;
            object one = 1;
            object chr = Word.WdUnits.wdCharacter;
            int resultRng = 0;
            int startRng;
            Word.Range TmpRng;

            IEnumerator paraEnum = wrkRng.Paragraphs.GetEnumerator();
            while (paraEnum.MoveNext())
            {
                Word.Paragraph para = (Word.Paragraph)paraEnum.Current;
               // System.Windows.Forms.MessageBox.Show("Whole Rnage  " + para.Range.Start + " === " + para.Range.End);

                TmpRng = para.Range.Duplicate;                  
                startRng = para.Range.Start;
                resultRng = 0;  //Resetting it for each paragraph.

                while (resultRng != -1)
                {
                    
                     resultRng = GetManualLineBreakRange(TmpRng);

                 //   System.Windows.Forms.MessageBox.Show(TmpRng.Start + " -- " + TmpRng.End + " === " + resultRng + "~~~" + TmpRng.Text);

                    if (resultRng != -1)
                    {
                        TmpRng.SetRange(startRng, resultRng);

                        str = TmpRng.Text.Trim();
                        cleanStr = StripFTFormats(str);
                        if (cleanStr.StartsWith("∙") || cleanStr.StartsWith("•") || cleanStr.StartsWith(""))
                        {
                            //System.Windows.Forms.MessageBox.Show(TmpRng.Start + " -- " + TmpRng.End + " === " + resultRng + "~~~" + TmpRng.Text);
                            setStyle(tspdDoc_, "List Bullet 2", TmpRng);
                            TmpRng.Collapse(ref WordHelper.COLLAPSE_END);
                        }
                        else if (str.StartsWith("▫") || str.StartsWith("o"))
                        {
                            //System.Windows.Forms.MessageBox.Show(TmpRng.Start + " -- " + TmpRng.End + " === " + resultRng + "~~~" + TmpRng.Text);
                            setStyle(tspdDoc_, "List Bullet 3", TmpRng);
                            TmpRng.Collapse(ref WordHelper.COLLAPSE_END);
                        }
                        else
                        {
                            //para.Range.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorDarkRed;
                        }

                        startRng = TmpRng.End;
                        TmpRng.SetRange(startRng, para.Range.End);

                    }
                    else
                    {
                    // TmpRng.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlue;
                    }
                }

                //str = para.Range.Text.Trim();
                //cleanStr = StripFTFormats(str);

                //if (cleanStr.StartsWith("∙") || cleanStr.StartsWith("•") || cleanStr.StartsWith(""))
                //{
                //    setStyle(tspdDoc_, "List Bullet 2", para.Range);
                //    para.Range.Collapse(ref WordHelper.COLLAPSE_END);
                //}
                //else if (str.StartsWith("▫") || str.StartsWith("o"))
                //{
                //    setStyle(tspdDoc_, "List Bullet 3", para.Range);
                //    para.Range.Collapse(ref WordHelper.COLLAPSE_END);
                //}
                newEndRng = para.Range.End;
            }

            wrkRng.SetRange(wrkRng.Start, newEndRng);
        }

        public static int GetManualLineBreakRange(Word.Range _selRng)
        {
            try
            {
                bool found = false;
                object _optMissing = System.Reflection.Missing.Value;

                Word.Range sel_ = _selRng;
                sel_.Find.Text = "^l";
                sel_.Find.Forward = true;
                sel_.Find.Wrap = Word.WdFindWrap.wdFindStop;
                sel_.Find.MatchCase = false;
                sel_.Find.MatchWholeWord = false;
                sel_.Find.MatchWildcards = false;
                

                found = sel_.Find.Execute(
                      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing,
                      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing,
                      ref _optMissing, ref _optMissing, ref _optMissing,
                      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing);

                if (found)
                {
                    return sel_.End;
                }
                else
                {
                    return -1;
                }

            }
            catch (Exception ex)
            {

            }

            return -1;
        }


    }
}