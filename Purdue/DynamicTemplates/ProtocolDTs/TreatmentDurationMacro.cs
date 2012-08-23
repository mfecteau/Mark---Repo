using System;
using System.Collections;
using System.Windows.Forms;

using Tspd.Tspddoc;
using Tspd.MacroBase;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;
using MSXML2;

using Word = Microsoft.Office.Interop.Word;

namespace VersionControl 
{
    internal sealed class TreatmentDurationMacro
	{
        private static readonly string header_ = @"$Header: Treatment.cs, 1, 50-jul-10 11:05:10, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
    /// Summary description for DurationofTreatmentMacro.
	/// </summary>
	public class TreatmentDurationMacro : AbstractMacroImpl
	{
        public TreatmentDurationMacro(MacroExecutor.MacroParameters mp)
            : base(mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods

        #region TreatmentDurationMacro
        /// <summary>
        /// /// Displays contact information (Fax only) based on Role Type
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
        public static MacroExecutor.MacroRetCd ShowDuration(
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.TreatmentDurationMacro.ShowDuration,ProtocolDTs.dll" elementLabel="TimesByTask" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Schedule" autogenerates="true" toolTip="TimesByTask" shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>
#endif
            try 
			{
                mp.pba_.setOperation("Duration of Treatment Macro", "Generating information...");
                TreatmentDurationMacro macro = null;

                macro = new TreatmentDurationMacro(mp);
				macro.preProcess();
				macro.display();
              //  macro.displayFax();
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
                Log.exception(e, "Error in Duration of Treatment Macro");
                mp.inoutRng_.Text = "Duration of Treatment Macro: " + e.Message;
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
        public MacrosConfig mc = null;
        public SOA _currentSOA = null;  
        public long totalWeeks = 0;

		public override void display()
		{
            string msg = "";
            Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;

			pba_.updateProgress(1.0);


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
                return;
            }

            if (_currentSOA == null)
            {
                //SOA is NULL
                msg = mc.getMessageByName("exception1").Text;
                wrkRng.InsertAfter(msg);
                wrkRng.InsertParagraphAfter();
                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                // Set outgoing range
                inoutRange.End = wrkRng.End;
                setOutgoingRng(inoutRange);
                wdDoc_.UndoClear();
                return;
            }


            PeriodEnumerator peEnum = _currentSOA.getPeriodEnumerator();

            if (peEnum.getList().Count<=0)
            {
                //No Periods are defined
                msg = mc.getMessageByName("exception2").Text;
                wrkRng.InsertAfter(msg);
                wrkRng.InsertParagraphAfter();
                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                // Set outgoing range
                inoutRange.End = wrkRng.End;
                setOutgoingRng(inoutRange);
                wdDoc_.UndoClear();
                return;
            }

            string strPer = "";
            string strsubPer = "";
            string strMain = "";
            string openBracket = mc.getMessageByName("symbol1").Text;  // "("
            string closeBracket = mc.getMessageByName("symbol2").Text;  // ")"
            string delimeter = mc.getMessageByName("symbol3").Text; //  ";"
            string lastperiod = mc.getMessageByName("symbol4").Text; //  ";"
            string strNull = mc.getMessageByName("null").Text;  //null

            strPer = mc.getMessageByName("period").Text;
            strsubPer = mc.getMessageByName("subperiod").Text;

            string tmpPer = "";
            string tmpSubPer = "";
            ArrayList arrSubper = new ArrayList();
            long strRetResult = 0;

            FTSpan objTimeSpan = null;

            foreach(Period p in peEnum.getList())
            {
                //Clear local var for each period
                tmpPer = "";
                tmpSubPer = "";

                //Handle Period
                if (LittleUtilities.isEmpty(p.getDuration()))
                {
                    totalWeeks = totalWeeks + 0;
                    strRetResult = 0;
                }
                else
                {
                    //ConvertToWeeks(p.getDuration(), p.getDurationTimeUnit());
                    totalWeeks += Convert.ToInt64(p.getDuration());

                    objTimeSpan = new FTSpan(Convert.ToInt64(p.getDuration()), FTSpan.DEFAULT_TIME_UNIT);
                    strRetResult = objTimeSpan.getTotalValue(p.getDurationTimeUnit());

                }


                tmpPer = strPer.Replace("[[period]]", p.getActualDisplayValue());
                tmpPer = tmpPer.Replace("[[duration]]", strRetResult.ToString());
                tmpPer = tmpPer.Replace("[[unit]]", p.getDurationTimeUnit());

                arrSubper.Clear(); //Clearing arrayList

                EventScheduleEnumerator subPerChildren = _currentSOA.getPeriodChildren(p);                
                foreach (EventScheduleBase subPrd in subPerChildren.getList())
                {
                      try
                    {
                        Period p1 = (Period)subPrd;   /// Just making sure, Visits are not included.
                        if (p1.isSubPeriod())
                        {
                         tmpSubPer = strsubPer.Replace("[[subperiod]]", p1.getActualDisplayValue());
                         if (LittleUtilities.isEmpty(p1.getDuration()))
                         {
                             tmpSubPer = tmpSubPer.Replace("[[duration]]", strNull);
                             tmpSubPer = tmpSubPer.Replace("[[unit]]", "");
                         }
                         else 
                         {

                             objTimeSpan = new FTSpan(Convert.ToInt64(p1.getDuration()), FTSpan.DEFAULT_TIME_UNIT);

                            strRetResult = objTimeSpan.getTotalValue(p1.getDurationTimeUnit());
                            tmpSubPer = tmpSubPer.Replace("[[duration]]", strRetResult.ToString());
                             tmpSubPer = tmpSubPer.Replace("[[unit]]", p1.getDurationTimeUnit());
                         }
                        

                            //For formatting purpose, please add it to an arraylist
                            arrSubper.Add(tmpSubPer.Trim());
                        }
                    }
                    catch (Exception ex)
                    {
                        //Ignore the exception, as we know Visits cannot be type casted as protocol.
                    }
                } //End For


                if (arrSubper.Count > 0)
                {
                    tmpSubPer = "";  //reseting it back as they are stored in arraylist.
                    for (int i = 0; i < arrSubper.Count; i++)
                    {
                        tmpSubPer += arrSubper[i];
                        if (arrSubper.Count > 1)
                        {
                            if (i + 1 == arrSubper.Count - 1) //As Its ZERO based index
                            {
                                tmpSubPer += " and ";
                            }
                            else
                            {
                                if (i != arrSubper.Count - 1)
                                {
                                    tmpSubPer += ", ";
                                }
                            }
                        }
                    }

                    strMain += tmpPer + openBracket + tmpSubPer + closeBracket + delimeter + " ";
                }
                else
                {
                    //If no Sub Period
                    strMain += tmpPer + delimeter + " ";
                }
            } //End Period For Loop

            
            strMain = strMain.Trim().Substring(0, strMain.Trim().Length - 1) + lastperiod;

            ///Start Printing.
            ///

            mc.setStyle(mc.getMessageByName("firstline").Format.Style, tspdDoc_, wrkRng);
            msg = mc.getMessageByName("firstline").Text;
            msg = msg.Replace("[[total]]", ConvertToWeeks(totalWeeks));
            WordFormatter.FTToWordFormat2(ref wrkRng, msg);
            wrkRng.InsertParagraphAfter();
            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

            WordFormatter.FTToWordFormat2(ref wrkRng, strMain);
            wrkRng.InsertParagraphAfter();
            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			// Set outgoing range
			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();
		}

        private string ConvertToWeeks(long duration)
        {
            //This function converts all time units into Week. It accepts time unit as "hours, days, Years & Months"
            // duration will bein string format, but it will always be NUMERIC (FRONT END)

            double _ans =0;
            try
            {

                long _valinWeeks = duration / (7 * 24 * 60 * 60);  //(7 days, 24 hours, 60mins, 60 sec)

               _ans =  Math.Round(Convert.ToDouble(_valinWeeks));

               if (_ans < 1.0)
               {
                   return "0";
               }
            }
            catch (Exception e)
            {
                Log.exception(e, e.Message + " - Converting to weeks.");
            }
            return _ans.ToString();
        }


        public override void preProcess()
        {

            string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);
            if (MacroBaseUtilities.isEmpty(elementPath))
            {
                return;
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
        }

		public override void postProcess()
		{
			// Clean up memory
		}
	}
}
