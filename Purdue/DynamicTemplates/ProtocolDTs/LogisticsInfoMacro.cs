using System;
using System.Collections;

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
	internal sealed class LogisticsInfoMacro
	{
		private static readonly string header_ = @"$Header: LogisticsInfoMacro.cs, 1, 18-Aug-09 12:04:44, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for Cohort Info.
	/// It basically calculates the information for each Arm/Cohorts.
	/// </summary>
	public class LogisticsInfoMacro : AbstractMacroImpl
	{
		
		SOA soa_ = null;
		bool DISPLAY_OTHER = false;

		public LogisticsInfoMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
			
		#region LogisticsInfoMacro


		public static MacroExecutor.MacroRetCd OtherCost (MacroExecutor.MacroParameters mp) 
		{

			try 
			{
				mp.pba_.setOperation("Calculating Cost Information", "Generating information...");
				
				LogisticsInfoMacro macro = null;
				macro = new LogisticsInfoMacro(mp);
				macro.preProcess();
				macro.DISPLAY_OTHER = true;
				macro.display();

				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Other Cost Information Macro"); 
				mp.inoutRng_.Text = "Other Cost Information Macro: " + e.Message;
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
//			if (MacroBaseUtilities.isEmpty(elementPath)) 
//			{
//				return;
//			}

			SOAEnumerator soaEnum = bom_.getAllSchedules();
				
			while (soaEnum.MoveNext())
			{
				pba_.updateProgress(2.0);

				SOA soa = soaEnum.getCurrent();
					soa_ = soa;
					break;
				
			}

			if (soa_ == null) return;
			
		}





		public override void display()
		{
			if (DISPLAY_OTHER)
			{
				DisplayOtherCost();
		
				
			}
		}
		public void DisplayOtherCost()
		{
			Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;

			pba_.updateProgress(1.0);
			
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);			

            wrkRng.InsertParagraphAfter();
	
			wrkRng.Text = getMedCosts() + " per subject";
			wrkRng.ParagraphFormat.LineSpacingRule= (Word.WdLineSpacing.wdLineSpace1pt5);
			
//			Word.Table table =  createTable(ref wrkRng, 3, 2);
//
//			Word.Row row = null;
//			Word.Cell cell = null;
//			Word.Cell cell1 = null;
//			Word.Cell cell2 = null;
//
//			row = table.Rows.Item(1);
//
//			cell1 = row.Cells.Item(1);
//			cell1.Range.Text = "LOGISTICS";
//
//			cell2 =row.Cells.Item(2);
//			cell2.Range.Text="";
//
//			cell1.Merge(cell2);
//
//
//			row = table.Rows.Item(2);
//
//			cell1 = row.Cells.Item(1);
//			cell1.Range.Text = "Estimated Investigator Grant Costs $ ";
//
//			cell2 =row.Cells.Item(2);
//			cell2.Range.Text = getMedCosts() + " [cost of single cohort]";
//
//			
//
//			row = table.Rows.Item(3);
//
//			cell1 = row.Cells.Item(1);
//			cell1.Range.Text = "Other External Costs (central Lab etc.) $";
//
//			cell2 =row.Cells.Item(2);
//			cell2.Range.Text="";
//
//			

//			wrkRng.End= MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);			

			// Set outgoing range
			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();

		}

		private string getMedCosts() 
		{
			if (soa_ == null) 
			{
				return "";
			}
			
			TaskEnumerator ta = soa_.getTaskEnumerator();
			double sumCosts = 0;

			while (ta.MoveNext())
			{
				Task task = ta.Current as Task;
				// Get the cost of the task
				double cost = getCostForTask(task);
				sumCosts += cost;
				
			}

			ArmEnumerator ae = bom_.getArmEnumerator();
			int count = ae.getList().Count;
			long subs = 1;
			if(count > 0)
			{
				ae.MoveNext();
				Arm a = ae.Current as Arm;
				try 
				{
					subs =  Convert.ToInt32(a.getPlannedEnrollmentPerArm());
				} 
				catch (Exception ex) {}
			}
			sumCosts *= subs;

			
			/*
						CurrencyUtilities currUtils = BridgeProxy.getInstance().getCurrencyUtilities(
							getTspdDocument().getDocumentDetails().getUser());
						FTLong currID = currUtils.getCurrencyID(); 

						sumCosts = currUtils.convertFromUSToLocal(sumCosts);
						String text = currUtils.getCurrencySymbol(_currID) + " " +
							CurrencyUtilities.roundUp(cost,0);
					}}

				sumCosts = currUtils.convertFromUSToLocal(sumCosts);
				lblTotalCPPValue.Text = currUtils.getCurrencySymbol(_currID) + " " +
				CurrencyUtilities.roundUp(sumCosts, 0);*/

			return ("(" + subs + " subjects) $ " + CurrencyUtilities.roundUp(sumCosts, 0));

		}

		private double getCostForTask(Task task) 
		{
			double cost = 0.0;
			TaskVisitEnumerator tvEnum = soa_.getTaskVisitsForTask(task);
			while (tvEnum.MoveNext())
			{
				TaskVisit tv = tvEnum.Current as TaskVisit; 
				//if (_allVisits.ContainsKey(tv.getAssociatedVisitID())) 
			{
				cost += task.getCost();
			}
			}

			return cost;

		}
		private void setTableBodyFont(ref Word.Font targetFont, SOATableFormat tblFmt) 
		{
			if (tblFmt.getDocTableBodyFontName() != null)
			{
				targetFont.Name = tblFmt.getDocTableBodyFontName();
			}

			if (tblFmt.getDocTableBodyFontSize() != null)
			{
				targetFont.Size = tblFmt.getDocTableBodyFontSize();
			}

			targetFont.Bold = (tblFmt.getDocTableBodyFontBold() ? -1 : 0);	
			targetFont.Italic = (tblFmt.getDocTableBodyFontItalics() ? -1 : 0);
		}

		private Word.Table createTable(ref Word.Range wrkRng, int rows, int cols) 
		{
			// Turn off auto caption for Word tables.
			Word.AutoCaption ac = wdApp_.AutoCaptions.get_Item(ref WordHelper.AUTO_CAPTION_WORD_TABLE);
			bool oldState = ac.AutoInsert;
			ac.AutoInsert = false;

			if (false) 
			{
				object tableCaption = "\ttable caption";
				wrkRng.InsertAfter(" "); // single space which'll get shifted after the caption.
				Word.Range captionRng = wrkRng.Duplicate;
				captionRng.Collapse(ref WordHelper.COLLAPSE_START);

				// put the caption in if the TableView gives you one...
				if (tableCaption != null)
				{
					captionRng.InsertCaption(
						ref WordHelper.CAPTION_LABEL_TABLE, ref tableCaption,
                        ref VBAHelper.OPT_MISSING, ref WordHelper.CAPTION_POSITION_ABOVE, ref VBAHelper.OPT_MISSING);
				}

				// Mark the added single space and replace it with a paragraph mark.
				wrkRng.Start = wrkRng.End - 1;
				wrkRng.InsertParagraph();

				// Collapse the range to the end of the paragraph mark so that the table can be added
				// after it.
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}

			// Insert the table. Note, that the table is inserted starting at but after the range.
			// So viewRng isn't increased.

			Word.Table tbl = wdDoc_.Tables.Add(
				wrkRng, rows, cols,
				ref WordHelper.WORD8_TABLE_BEHAVIOR, ref VBAHelper.OPT_MISSING);

			// Reinstate auto caption for Word tables.
			ac.AutoInsert = oldState;


			// Autofit and table sizing are problematic, use the monkey
			tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);

			tbl.Borders.Enable = VBAHelper.iTRUE;
			tbl.TopPadding = 0;
			tbl.BottomPadding = 0;
			tbl.LeftPadding = 0;
			tbl.RightPadding = 0;
			tbl.Spacing = 0;


			tbl.Columns[1].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
			tbl.Columns[1].PreferredWidth = 75.0f;
			tbl.Columns[2].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
			tbl.Columns[2].PreferredWidth = 25.0f;
			tbl.Rows.LeftIndent = tbl.Application.InchesToPoints(0f);

			tbl.Rows.AllowBreakAcrossPages = VBAHelper.iFALSE;

			tbl.Borders.OutsideLineWidth = MacroBaseUtilities.LINE_WIDTH_NORMAL;
			tbl.Borders.InsideLineWidth = MacroBaseUtilities.LINE_WIDTH_NORMAL;

			wrkRng.End = tbl.Range.End;
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			wdDoc_.UndoClear();

			return tbl;
		}

		private Word.Table createFootnoteTable(Word.Range viewRng, int numFootnotes) 
		{
			// Turn off auto caption for Word tables.
			Word.AutoCaption ac =
				wdApp_.AutoCaptions.get_Item(ref WordHelper.AUTO_CAPTION_WORD_TABLE);
			bool oldState = ac.AutoInsert;
			ac.AutoInsert = false;

			Word.Range wrkRng = viewRng.Duplicate;

			// Collapse the range to the end of the paragraph mark so that the table can be added
			// after it.
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			wrkRng.InsertParagraphAfter();
			object oStyle = tspdDoc_.getStyleHelper().setNamedStyle(PfizerUtil.PFIZER_STYLE_TABLETEXT_10, wrkRng);
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			int nbrCols = 2;
			int nbrRows = numFootnotes;

			// Insert the table. Note, that the table is inserted starting at but after the range.
			// So viewRng isn't increased.
			Word.Table tbl =
				wdDoc_.Tables.Add(
				wrkRng, nbrRows, nbrCols,
				ref WordHelper.WORD8_TABLE_BEHAVIOR, ref VBAHelper.OPT_MISSING);

			oStyle = tspdDoc_.getStyleHelper().setNamedStyle(PfizerUtil.PFIZER_STYLE_TABLETEXT_10, tbl.Range);

			// Reinstate auto caption for Word tables.
			ac.AutoInsert = oldState;

			tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);

			tbl.Borders.Enable = VBAHelper.iFALSE;
			tbl.LeftPadding = 0;
			tbl.RightPadding = 0;
			tbl.Spacing = 0;
			tbl.Rows.LeftIndent = tbl.Application.InchesToPoints(0f);

			tbl.Rows.AllowBreakAcrossPages = VBAHelper.iFALSE;

			// Increase viewRng to include the table.
			viewRng.End = tbl.Range.End;

			wdDoc_.UndoClear();

			return tbl;
		}

		
	
		public override void postProcess()
		{
			// Clean up memory
		
		}
	}
}
