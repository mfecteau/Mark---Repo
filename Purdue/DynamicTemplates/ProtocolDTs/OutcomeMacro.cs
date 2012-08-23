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
	internal sealed class OutcomeMacro
	{
		private static readonly string header_ = @"$Header: OutcomeMacro.cs, 1, 18-Aug-09 12:05:03, Pinal Patel$";
	}
}


namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for ObjectiveMacro.
	/// </summary>
	public class OutcomeMacro : AbstractMacroImpl
	{
		public string outcomeType;
		public string outcomeLabel;
		private ArrayList outcomes = new ArrayList();
        public string _SelectedType = "";
        public string sel_Type = "";
        public MacrosConfig mc = null;

		public OutcomeMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods

		public static MacroExecutor.MacroRetCd Outcomes(
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.SalesDemo.DynTmplts.OutcomeMacro.Outcomes,ProtocolDTs.dll" elementLabel="Outcome" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Outcomes" autogenerates="true" toolTip="Lists outcomes." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Outcome Macro", "Generating information...");				
				OutcomeMacro macro = null;
				macro = new OutcomeMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in  Outcome Macro"); 
				mp.inoutRng_.Text = "Outcome Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}		
		
		#endregion


		public override void preProcess() 
		{
			try 
			{			
				

			} 
			catch (Exception e) 
			{
				Log.exception(e, "Problem in preprocess()");
				throw e;
			}
		}


		public override void display() 
		{
			Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;

			pba_.updateProgress(1.0);
       
			string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);


            string chooserElementPath = this.macroEntry_.getElementPath();
            string fPath = tspdDoc_.getTrialProject().getTemplateDirPath() + "\\dyntmplts\\MacrosConfig.xml";
            mc = new MacrosConfig(fPath, chooserElementPath);


			bool isOther;

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
						sel_Type = aParms[0];
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
						outcomeLabel = aParms[1];
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
                        _SelectedType = aParms[2];
                    }
                    catch (Exception ex)
                    {
                        parmsValid = false;
                    }
                }
			}


            bool doNothing = false;
				// Ask the user if the parms are missing/invalid
				if (!parmsValid) 
				{
					outcomes.Clear();					
					/*****/

					OutcomeSelection frmSel = new OutcomeSelection();
                    frmSel.Text = mc.getMessageByName("captiontext").Text;
                    _SelectedType = mc.getMessageByName("selectiontype").Text;
                    frmSel.label1.Text = mc.getMessageByName("label1").Text;

                    if (_SelectedType.ToLower() == "optional")
                    {
                        frmSel.rdbyOutcome.Text = mc.getMessageByName("rdoutcome").Text;
                        frmSel.rdbyObjective.Text = mc.getMessageByName("rdobjective").Text;
                        frmSel.outcomeTooltip.SetToolTip(frmSel.rdbyOutcome, mc.getMessageByName("tooltip1").Text);
                        frmSel.outcomeTooltip.SetToolTip(frmSel.rdbyObjective, mc.getMessageByName("tooltip2").Text);
                    }

                    frmSel.LoadOutcomes(bom_,_SelectedType);

                    if (frmSel.DialogResult == System.Windows.Forms.DialogResult.OK)
                    {
                        sel_Type = frmSel.var_Type;
                        outcomeLabel = frmSel.varLabel;

                        if (_SelectedType.ToLower() == "optional")
                        {
                            if (frmSel.rdbyOutcome.Checked)
                            {
                                _SelectedType = "outcometype";
                            }
                            else if (frmSel.rdbyObjective.Checked)
                            {
                                _SelectedType = "objectivetype";
                            }
 
                        }

                        if (sel_Type != "FT_NA")
                        {
                            if (_SelectedType.ToLower() == "outcometype")
                            {
                                LoadOutcomes();
                            }
                            else if (_SelectedType.ToLower() == "objectivetype")
                            {
                                LoadObjectives();
                            }
                            execParms_.setParm(MacroExecutor.MacroExecParameters.PARM_1, sel_Type + "|" + outcomeLabel + "|" + _SelectedType);
                        }
                        else
                        {
                            execParms_.setParm(MacroExecutor.MacroExecParameters.PARM_1, sel_Type + "|" + outcomeLabel + "|" + _SelectedType);
                            wrkRng.InsertAfter(mc.getMessageByName("exception2").Text);
                            inoutRange.End = wrkRng.End;
                            setOutgoingRng(inoutRange);
                            wdDoc_.UndoClear();
                            return;
                        }
                        //execParms_.setParm(MacroExecutor.MacroExecParameters.PARM_1, sel_Type + "|" + outcomeLabel + "|" + _SelectedType);
                    }
                    else  //If dialog result is Cancel.
                    {
                        doNothing = true;
                        pba_.done();
                      // this.macroStatusCode = MacroExecutor.MacroRetCd.Cancelled;
                       this.MacroStatusCode = MacroExecutor.MacroRetCd.Cancelled;
                       return;
                    }
				}
				else
				{
                   
                        if (_SelectedType.ToLower() == "outcometype")
                        {
                            LoadOutcomes();
                        }
                        else if (_SelectedType.ToLower() == "objectivetype")
                        {
                            LoadObjectives();
                        }
				}

                
                    displayOutcome(ref wrkRng);
                    mc.setStyle(mc.getMessageByName("normalstyle").Format.Style, tspdDoc_, wrkRng);
                

			// Set outgoing range
			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();
		}

        private void LoadObjectives()
        {
            /** This method gets Outcomes based on Objective Type.
             * ***/


            OutcomeEnumerator oe = bom_.getOutcomes();
            int count = bom_.getOutcomes().getList().Count;
            ArrayList arrObjID = new ArrayList();

            double progInc = 30.0 / (double)count;

            while (oe.MoveNext())
            {
                Outcome outcome1 = (Outcome)oe.Current;
                ObjectiveEnumerator objEnum = bom_.getAssociatedObjectives(outcome1);

                while (objEnum.MoveNext())
                {
                    Objective PriObj = (Objective)objEnum.Current;

                        if (PriObj.getObjectiveType().ToLower().Equals(sel_Type.ToLower()))
                        {
                            if (PriObj.getObjectiveType().ToLower() == "other")
                            {
                                if (PriObj.getOtherObjective() == outcomeLabel)
                                {
                                    if (arrObjID.Contains(outcome1.getObjID().ToString()) == false)
                                    {
                                        outcomes.Add(outcome1);
                                        arrObjID.Add(outcome1.getObjID().ToString());
                                    }

                                }
                            }
                            else
                            {
                                if (arrObjID.Contains(outcome1.getObjID().ToString()) == false)
                                {
                                    outcomes.Add(outcome1);
                                    arrObjID.Add(outcome1.getObjID().ToString());
                                }
                            }
                           
                        }
                    
                }
                pba_.updateProgress(progInc);
            }
        }

        private void LoadOutcomes()
        {
            OutcomeEnumerator oe = bom_.getOutcomes();
            int count = bom_.getOutcomes().getList().Count;

            double progInc = 30.0 / (double)count;

            while (oe.MoveNext())
            {
                Outcome outcome1 = (Outcome)oe.Current;
                if (outcome1.getOutcomeType().ToLower().ToString() == sel_Type.ToLower())
                {
                    if (sel_Type == "other")
                    {
                        if (outcome1.getOtherOutcome() == outcomeLabel)
                        {
                            outcomes.Add(outcome1);
                        }
                    }
                    else
                    {
                        outcomes.Add(outcome1);
                    }
                }
            }	
 
        }

		private void displayOutcome(ref Word.Range wrkRng)
		{
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
            string msg = "";
			

			if (outcomes.Count == 0) 
			{
                msg = mc.getMessageByName("exception1").Text;
                msg = msg.Replace("[[outcome]]",outcomeLabel);
				wrkRng.InsertAfter(msg);
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				return;
			}

            int start = wrkRng.Start;


            MacrosConfig.message msg1 = null;
            msg1 = mc.getMessageByName("firstline");
            if (msg1.Text.Length > 0)
            {
             //   mc.setFirstLineStyle(tspdDoc_, wrkRng);
                msg = msg1.Text;
                msg = msg.Replace("[[outcome]]", outcomeLabel);
                wrkRng.InsertAfter(msg);
                mc.setStyle(msg1.Format.Style, tspdDoc_, wrkRng);
                //wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                
            }
			//wrkRng.InsertAfter(outcomeLabel + " outcome(s) ");
			Word.Range rngPrimary = wrkRng.Duplicate;
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			//rngPrimary.Font.Bold = VBAHelper.iTRUE;

			double progInc = 20.0 / (double)outcomes.Count;

			
            mc.setStyle(mc.getMessageByName("bodytext").Format.Style, tspdDoc_, wrkRng);

			foreach (Outcome obj1 in outcomes)
			{
				pba_.updateProgress(progInc);
                msg = mc.getMessageByName("bodytext").Text;
                msg = msg.Replace("[[outcome]]",obj1.getFullDescription());
                msg = msg.Replace("\n", "\v");
                msg = msg.Replace("\r", "\v");
                wrkRng.InsertAfter(msg);
                wrkRng.InsertParagraphAfter();
			}

            int end = wrkRng.End;
            wrkRng.SetRange(start, end);

            Tspd.Utilities.WordFormatter.FTToWordFormat2(ref wrkRng, wrkRng.Text);

			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			wdDoc_.UndoClear();
		}
		public override void postProcess()
		{
			// Clean up memory
			outcomes.Clear();
		}

	
	}
}
