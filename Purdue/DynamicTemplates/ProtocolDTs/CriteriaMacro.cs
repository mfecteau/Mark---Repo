using System;
using System.Collections;

using Tspd.Tspddoc;
using Tspd.MacroBase;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;
using MSXML2;

using TspdCfg.FastTrack.DynTmplts;

using Word = Microsoft.Office.Interop.Word;

namespace VersionControl 
{
	internal sealed class CriteriaMacro
	{
		private static readonly string header_ = @"$Header: CriteriaMacro.cs, 1, 18-Aug-09 12:03:32, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for CriteriaMacro.
	/// </summary>
	public class CriteriaMacro : AbstractMacroImpl
	{
		private ArrayList criteria = new ArrayList();
		
		private string crit_type_;
        private string crit_lablel_;
		private string listStyle_;
		private string headingStyle_;

        MacrosConfig mc = null;

		public CriteriaMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region InclusionCriteria
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd Criteria (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.CriteriaMacro.InclusionCriteria,ProtocolDTs.dll" elementLabel="Inclusion Criteria" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Eligibility Criteria" autogenerates="true" toolTip="Lists inclusion criteria." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Inclusion Criteria Macro", "Generating information...");
				
				CriteriaMacro macro = null;
				macro = new CriteriaMacro(mp);

				//macro.type_ = PfizerUtil.INCLUSION;
                macro.listStyle_ = PurdueUtil.PFIZER_INC_CRIT_NUMBERED_LISTS;
                //macro.headingStyle_ = PurdueUtil.PFIZER_STYLE_TEXT_TI12_LEFT;
				
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Inclusion Criteria Macro"); 
				mp.inoutRng_.Text = "Inclusion Criteria Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		#endregion



		#region ExclusionCriteria
		/// <summary>
		/// Displays all exclusion criteria with category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd ExclusionCriteria (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.CriteriaMacro.ExclusionCriteria,ProtocolDTs.dll" elementLabel="Exclusion Criteria" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Eligibility Criteria" autogenerates="true" toolTip="Lists exclusion criteria." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Exclusion Criteria Macro", "Generating information...");
				
				CriteriaMacro macro = null;
				macro = new CriteriaMacro(mp);

				//macro.type_ = PfizerUtil.EXCLUSION;
				macro.listStyle_ = PurdueUtil.PFIZER_INC_CRIT_NUMBERED_LISTS;
				

				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Exclusion Criteria Macro"); 
				mp.inoutRng_.Text = "Exclusion Criteria Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		#endregion



		#endregion


        public override void preProcess()
        {
            try
            {
                string chooserElementPath = this.macroEntry_.getElementPath();
                string fPath = tspdDoc_.getTrialProject().getTemplateDirPath() + "\\dyntmplts\\MacrosConfig.xml";
                mc = new MacrosConfig(fPath, chooserElementPath);
                criteria.Clear();
            }
            catch (Exception e)
            {
                Log.exception(e,e.Message);
                throw e;
            }
        }

		public override void display() 
		{
            string subtype = "";
            ArrayList critSubtype = new ArrayList();

            // Get stored parameters
            string sParms = execParms_.getParm(MacroExecutor.MacroExecParameters.PARM_1);
            string[] aParms = null;

            if (!MacroBaseUtilities.isEmpty(sParms))
            {
                aParms = sParms.Split('|');
            }

            bool parmsValid = false;

            if (aParms != null && aParms.Length == 3)
            {
                parmsValid = true;

                if (!MacroBaseUtilities.isEmpty(aParms[0]))
                {
                    try
                    {
                        crit_type_ = aParms[0];
                    }
                    catch (Exception ex)
                    {
                        parmsValid = false;
                    }
                }
                if (!MacroBaseUtilities.isEmpty(aParms[1]))
                {
                    try
                    {
                        crit_lablel_ = aParms[1];
                    }
                    catch (Exception ex)
                    {
                        parmsValid = false;
                    }
                }
                if (!MacroBaseUtilities.isEmpty(aParms[2]))
                {
                    try
                    {
                        subtype = aParms[2];
                        string[] stored_subtype = null;
                        stored_subtype = subtype.Split('@');

                        for (int k = 0; k < stored_subtype.Length; k++)
                        {
                            critSubtype.Add(stored_subtype[k]);
                        }
                    }
                    catch (Exception ex)
                    {
                        parmsValid = false;
                    }
                }
            }

           // System.Windows.Forms.CheckedListBox.CheckedItemCollection critSubtype = null;

            bool showUnclassified = false;
            Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;
			pba_.updateProgress(1.0);

            string label1 = mc.getMessageByName("label1").Text;
            string label2 = mc.getMessageByName("label2").Text;
            bool flagSubtype = Convert.ToBoolean(mc.getMessageByName("subtypevisible").Text);
            bool flagGrouping = Convert.ToBoolean(mc.getMessageByName("grouping").Text);
            //bool flagBullets = Convert.ToBoolean(mc.getMessageByName("usebullets").Text);

            if (!parmsValid)
            {
                CriteriaSelection cs = new CriteriaSelection();
                cs.Text = mc.getMessageByName("captiontext").Text;
                cs.LoadCriteria(tspdDoc_, bom_, icpSchemaMgr_, flagSubtype);
                if (cs.DialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    crit_type_ = cs.var_Type;  //
                    crit_lablel_ = cs.varLabel;
                    if (crit_type_ == "FT_NA")
                    {
                        //If nothing is selected and OK is clicked.
                        wrkRng.InsertAfter(mc.getMessageByName("exception2").Text);
                        inoutRange.End = wrkRng.End;
                        setOutgoingRng(inoutRange);
                        wdDoc_.UndoClear();
                        return;
                    }

                    foreach (object itemchecked in cs.chkLstSubType.CheckedItems)
                    {
                        critSubtype.Add(itemchecked.ToString().ToLower()); //collecting all checked SubTypes.
                        subtype += itemchecked.ToString() + "@";
                    }

                    if (critSubtype.IndexOf("unclassified") > -1)
                    {
                        showUnclassified = true;
                    }

                    if (flagSubtype)
                    {
                        //If flag sub-type =TRUE, it's grouping is considered as mandatory, else it would defeat the purpose of allowing selection of sub-types.
                        DisplayGroupedCriteria(wrkRng,critSubtype, showUnclassified);
                    }
                    else
                    {
                        if (flagGrouping)
                        {
                            //Apply Grouping
                            DisplayGroupedCriteria(wrkRng,critSubtype, showUnclassified);
                        }
                        else
                        {
                            //Do not check grouping, as its assumed & mandatory to Group it.
                            //Sub Type & Grouping both are false.
                            displayCriteria(wrkRng);
                        }
                    }

                    execParms_.setParm(MacroExecutor.MacroExecParameters.PARM_1, crit_type_ + "|" + crit_lablel_ + "|" + subtype);
                }
                else
                {
                    pba_.done();
                    // this.macroStatusCode = MacroExecutor.MacroRetCd.Cancelled;
                    this.MacroStatusCode = MacroExecutor.MacroRetCd.Cancelled;
                    return;
                }
                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                cs.Dispose();
            } //End paramsvalid
            else
            {
                 flagSubtype = Convert.ToBoolean(mc.getMessageByName("subtypevisible").Text);
                 flagGrouping = Convert.ToBoolean(mc.getMessageByName("grouping").Text);
                 if (crit_type_ == "FT_NA")
                 {
                     //If nothing is selected and OK is clicked.
                     wrkRng.InsertAfter(mc.getMessageByName("exception2").Text);
                     inoutRange.End = wrkRng.End;
                     setOutgoingRng(inoutRange);
                     wdDoc_.UndoClear();
                     return;
                 }
                 if (critSubtype.IndexOf("unclassified") > -1)
                 {
                     showUnclassified = true;
                 }

                 if (flagSubtype)
                 {
                     //If flag sub-type =TRUE, it's grouping is considered as mandatory, else it would defeat the purpose of allowing selection of sub-types.
                     DisplayGroupedCriteria(wrkRng,critSubtype, showUnclassified);
                 }
                 else
                 {
                     if (flagGrouping)
                     {
                         //Apply Grouping
                         DisplayGroupedCriteria(wrkRng,critSubtype, showUnclassified);
                     }
                     else
                     {
                         //Do not check grouping, as its assumed & mandatory to Group it.
                         //Sub Type & Grouping both are false.
                         displayCriteria(wrkRng);
                     }
                 }
            }
            

			// Set outgoing range
			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();
		}

        private void displayCriteria(Word.Range wrkRng)
        {
            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

            ElementListHelpers elh = new ElementListHelpers(tspdDoc_);
            IList critList = elh.getLiveChooserEntryListForCriteria();
            IEnumerator critIter = critList.GetEnumerator();
            while (critIter.MoveNext())
            {
                Criterion crit = (Criterion)critIter.Current;
               if (crit.getCriterionType().Trim().ToLower() == "other")
                {
                    if (crit.getOtherCriterion() == crit_lablel_)
                    {
                        criteria.Add(crit);
                    } 
                }
               else if (crit.getCriterionType().Trim().ToLower() == crit_type_.ToLower())
               {
                   criteria.Add(crit);
               }
            } //End While

            string msg = null;
            if (criteria.Count == 0)
            {
                msg = mc.getMessageByName("exception1").Text;
                msg = msg.Replace("[[criteriatype]]", crit_type_);
                wrkRng.InsertAfter(msg);
                wdDoc_.UndoClear();
                return;
            }

            int startRng = wrkRng.Start;

            
            MacrosConfig.message msg1 = null;
            msg1 = mc.getMessageByName("firstline");
            if (msg1.Text.Length > 0)
            {
               // mc.setStyle(msg1.Format.Style, tspdDoc_, wrkRng);
                msg = msg1.Text;
                msg = msg.Replace("[[criteriatype]]", crit_type_);
                wrkRng.InsertAfter(msg);
                wrkRng.InsertParagraphAfter();
                mc.setStyle(msg1.Format.Style, tspdDoc_, wrkRng);
                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
            }
            
            double progInc = 20.0 / (double)criteria.Count;
            bool flagRestartNo = false;
                mc.setStyle(mc.getMessageByName("bodytext").Format.Style, tspdDoc_, wrkRng);
                if (!flagRestartNo)
                {
                    mc.RestartNumbering(wrkRng, flagRestartNo);
                    flagRestartNo = true;
                }

            foreach (Criterion crit1 in criteria)
            {
                pba_.updateProgress(progInc);
                msg = mc.getMessageByName("bodytext").Text;              
                   msg = msg.Replace("[[criteria]]", crit1.getFullDescription().Replace("\n","\v"));
                   msg = msg.Replace("\r", "\v");
                //msg = msg.Replace("[[criteria]]", crit1.getFullDescription());
                wrkRng.InsertAfter(msg);
                wrkRng.InsertParagraphAfter();     
                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                wdDoc_.UndoClear();
            }

            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

            int endRng = wrkRng.End;
            wrkRng.SetRange(startRng, endRng);
            Tspd.Utilities.WordFormatter.FTToWordFormat2(ref wrkRng, wrkRng.Text);

            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

            mc.setStyle(mc.getMessageByName("normalstyle").Format.Style, tspdDoc_, wrkRng);
            wdDoc_.UndoClear();
        }


        private void DisplayGroupedCriteria(Word.Range wrkRng,ArrayList selected_Subtype,bool Disp_Unclassified)
        {
            try
            {
                ArrayList subtypeList = new ArrayList();
                ArrayList sortedcritList = new ArrayList();
                ArrayList unclassifiedList = new ArrayList();
                MacrosConfig.message msg1 = null;
                string msg = "";
                subtypeList = bom_.getIcpSchemaMgr().getEnumPairs("EntranceCriterionClassifierTypes");
                ElementListHelpers elh = new ElementListHelpers(tspdDoc_);
                IList critList = elh.getLiveChooserEntryListForCriteria();
                IEnumerator critIter = critList.GetEnumerator();

                int start = wrkRng.Start;
                bool flagRestartNo = false;
                bool FirstLinePrinted = false;  //to see if First line is printed

                for (int i = 0; i < subtypeList.Count; i++)
                {
                    EnumPair ep = (EnumPair)subtypeList[i];
                    critIter.Reset(); 
                    while (critIter.MoveNext())
                    {
                        Criterion crit = (Criterion)critIter.Current;
                        if (crit.getCriterionType().ToLower() == crit_type_.ToLower())
                        {
                            if (crit.getClassifierType().ToLower() == ep.getSystemName().ToLower())
                            {
                                if (selected_Subtype.IndexOf(ep.getUserLabel().ToLower()) > -1)  //See if its Checked or not.
                                {
                                    sortedcritList.Add(crit.getFullDescription());
                                }
                            }
                        }
                    }

                    if (sortedcritList.Count > 0)
                    {
                        if (!FirstLinePrinted)
                        {
                            msg1 = mc.getMessageByName("firstline");
                            if (msg1.Text.Length > 0)
                            { 
                                msg = msg1.Text;
                                msg = msg.Replace("[[criteriatype]]", crit_type_);
                                wrkRng.InsertAfter(msg);
                                wrkRng.InsertParagraphAfter();
                                mc.setStyle(mc.getMessageByName("firstline").Format.Style, tspdDoc_, wrkRng);
                                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                            }
                            FirstLinePrinted = true; 
                        }


                        msg1 = mc.getMessageByName("groupheadertext");
                        if (msg1.Text.Length > 0)
                        {
                            msg = msg1.Text;
                            msg = msg.Replace("[[subcriteriatype]]", ep.getUserLabel());
                            wrkRng.InsertAfter(msg);
                            wrkRng.InsertParagraphAfter();
                            mc.setStyle(msg1.Format.Style, tspdDoc_, wrkRng);
                            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);                            
                        }
                        
                        for (int j = 0; j < sortedcritList.Count; j++)
                        {
                                msg = mc.getMessageByName("bodytext").Text;
                                msg = msg.Replace("[[criteria]]", sortedcritList[j].ToString().Replace("\n", "\v"));
                                msg = msg.Replace("\r", "\v");
                                wrkRng.InsertAfter(msg);
                                wrkRng.InsertParagraphAfter();
                                mc.setStyle(mc.getMessageByName("bodytext").Format.Style, tspdDoc_, wrkRng);
                                if (!flagRestartNo)
                                {
                                    mc.RestartNumbering(wrkRng, flagRestartNo);
                                    flagRestartNo = true;
                                }
                                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);                              
                                wdDoc_.UndoClear();
                        }
                    }
                    sortedcritList.Clear();
                } //End For

                if (selected_Subtype.IndexOf("unclassified") > -1)
                {
                    //Printing Unclassified or with "NO SUB-TYPE"
                    critIter.Reset();  //reseting to start counter from 1
                    while (critIter.MoveNext())
                    {
                        Criterion crit = (Criterion)critIter.Current;
                        if (crit.getCriterionType().ToLower() == crit_type_.ToLower())
                        {
                            if (crit.getClassifierType().Trim().Length == 0)
                            {
                                unclassifiedList.Add(crit.getFullDescription());
                            }
                        }
                    }
                }

                if (unclassifiedList.Count > 0)
                {
                    if (!FirstLinePrinted)
                    {
                        msg1 = mc.getMessageByName("firstline");
                        if (msg1.Text.Length > 0)
                        {                            
                            msg = msg1.Text;
                            msg = msg.Replace("[[criteriatype]]", crit_type_);
                            wrkRng.InsertAfter(msg);
                            wrkRng.InsertParagraphAfter();
                            mc.setStyle(msg1.Format.Style, tspdDoc_, wrkRng);
                            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                        }
                        FirstLinePrinted = true;
                    }

                    msg1 = mc.getMessageByName("noclassifier");
                    if (msg1.Text.Length > 0)
                    {
                       
                        msg = msg1.Text;                       
                        wrkRng.InsertAfter(msg);
                        wrkRng.InsertParagraphAfter();
                        mc.setStyle(msg1.Format.Style, tspdDoc_, wrkRng);
                        wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                    }

                    for (int j = 0; j < unclassifiedList.Count; j++)
                    {
                        msg = mc.getMessageByName("bodytext").Text;
                        msg = msg.Replace("[[criteria]]", unclassifiedList[j].ToString().Replace("\n", "\v"));
                        msg = msg.Replace("\r", "\v");
                        wrkRng.InsertAfter(msg);
                        wrkRng.InsertParagraphAfter();
                        mc.setStyle(mc.getMessageByName("bodytext").Format.Style, tspdDoc_, wrkRng);
                        if (!flagRestartNo)
                        {
                            mc.RestartNumbering(wrkRng, flagRestartNo);
                            flagRestartNo = true;
                        }

                        wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                        wdDoc_.UndoClear();
                    }
                }


                if (sortedcritList.Count == 0 && unclassifiedList.Count == 0 )
                {
                    if (!FirstLinePrinted)
                    {
                        msg = mc.getMessageByName("exception1").Text;
                        msg = msg.Replace("[[criteriatype]]", crit_type_);
                        wrkRng.InsertAfter(msg);
                        wrkRng.InsertParagraphAfter();
                        wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

                        wdDoc_.UndoClear();
                        return;
                    }
                }

                mc.setStyle(mc.getMessageByName("normalstyle").Format.Style, tspdDoc_, wrkRng);
                wrkRng.InsertParagraphAfter();
                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

                int end = wrkRng.End;

                wrkRng.SetRange(start, end);
                Tspd.Utilities.WordFormatter.FTToWordFormat2(ref wrkRng, wrkRng.Text);

                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                wdDoc_.UndoClear();
            }
            catch (Exception ex)
            {
               //
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }




		
		public override void postProcess()
		{
			// Clean up memory
			criteria.Clear();
		}
	}
}
