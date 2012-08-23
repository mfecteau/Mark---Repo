using System;
using System.Globalization;
using System.Collections;
using Tspd.MacroBase;
using Tspd.MacroBase.Table;
using Tspd.MacroBase.BaseImpl;
using Tspd.Macros;
using Tspd.Utilities;
using System.Reflection;
using Tspd.Businessobject;
using Word = Microsoft.Office.Interop.Word;

namespace TspdCfg.Purdue.DynTmplts.Table
{
	/// <summary>
	/// Summary description for RocheSOATableDisplayMgr.
	/// </summary>
	public class PurdueSOATableDisplayMgr : DefSOATableDisplayMgr
	{

		bool _isBlinded = false;

		public PurdueSOATableDisplayMgr(MacroExecutor.MacroParameters mp) : base(mp)
		{
		}
		
		public bool BlindedStudy
		{
			set
			{
				_isBlinded = value;
			}
			get
			{
				return _isBlinded;
			}
		}


		public override DefSOATableView newTableView()
		{
			PurdueSOATableView rv = new PurdueSOATableView();
			rv.BlindedStudy = _isBlinded;

			return rv;
		}

		public override TableBuilder getTableBuilder()
		{
			if (tableBuilder_ == null)
			{
				tableBuilder_ = new PurdueTableHelper(this); 
			}

			return tableBuilder_;
		}

		public override bool useTableCaption() 
		{
			return false;
		}

		public override bool useHeaderRows() 
		{
			return true;
		}

		public override bool addCaptionRow() 
		{
			return true;
		}

		public override int getRowsPerPage() 
		{
			return 0;
		}

		public override void preProcess()
		{
			try 
			{
				
				//Adding code to update the Table of Contents & Creation Date.
				CreationDate_TOC();



				// bom_.clearFootnoteCollection();

				// use reflection to lookup 
				Type bomType = bom_.GetType();
				MethodInfo clearFootnoteCollectionMethod = bomType.GetMethod("clearFootnoteCollection");
				if (clearFootnoteCollectionMethod != null) 
				{
					clearFootnoteCollectionMethod.Invoke(bom_, null);
				}

				
			}
			catch (Exception ex) 
			{
			}

			base.preProcess();
		}

        public override void display()
        {
            string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);
            if (MacroBaseUtilities.isEmpty(elementPath))
            {
                macroStatusCode_ = MacroExecutor.MacroRetCd.Failed;
                return;
            }

            try
            {
                Word.Range wrkRng = startAtBeginningOfParagraph();

                //keep with next
                wrkRng.Paragraphs.KeepWithNext = VBAHelper.iTRUE;
                //turn off auto numbering
                //wrkRng.ListFormat.RemoveNumbers(ref WordHelper.NUMBER_PARAGRAPH);

                Word.Range begRng = wrkRng.Duplicate;
                Word.Range viewRng;

                if (!preDisplay(wrkRng))
                {
                    //impls precondition failed... so just return.
                    return;
                }

                // Insert a leading paragraph mark to give room for the wavy
                // red lines to do its work.
                wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
                wrkRng.Start = wrkRng.End;

                //reset fwdNoteNumberStyleArabicootnote numbering


                //Commented this code as if err's when concept locks are applied and someone tries to Print SOA. Also the reaosn to bring ti locally.

                //wrkRng.Footnotes.NumberStyle = Word.WdNoteNumberStyle.wdNoteNumberStyleArabic;
                //wrkRng.Footnotes.NumberingRule = Word.WdNumberingRule.wdRestartSection;

                // Ripped from GF.
                // Do not create section breaks before, between, or after tables
                // by default. If the user wants them, they'll create them
                // themselves.

                // This indicates whether a leading page break currently exists.
                // We keep this knowledge because we always think of page breaks
                // occurring in pairs but we don't always create 2 each time.
                // So if we created 2, then the next time we just need to create
                // 1 because it'll be paired with ending one of prior pair.
                bool hasLeadingPageBreak = false;

                IList wordTableViews = this.getWordTableView();
                double progInc = 70.0 / (double)wordTableViews.Count; //really more like 95%
                for (int i = 0; i < wordTableViews.Count; i++)
                {
                    TableView tableView = wordTableViews[i] as TableView;

                    if (tableView.getLeadingPageBreak() != null)  //has one!
                    {
                        if (i != 0)
                        {
                            wrkRng.InsertBreak(ref WordHelper.PAGE_BREAK);
                            // wrkRng.InsertParagraphAfter();
                            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                        }

                        // Display table.
                        displayTable(wrkRng, tableView);
                        wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                    }
                    else
                    {
                        // If the table before isn't sandwiched between page
                        // breaks, then we begin this table on a new line.
                        if (i != 0)
                        {
                            if (((TableView)wordTableViews[i - 1]).getLeadingPageBreak() == null)
                            {
                                wrkRng.InsertParagraphAfter();
                                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                            }
                        }

                        // The next table wanting page breaks will have to
                        // create both the leading and the matching.
                        hasLeadingPageBreak = false;

                        // Display table.
                        displayTable(wrkRng, tableView);

                        wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                    }

                    wdDoc_.UndoClear();

                    //update status bar
                    pba_.updateProgress(progInc);
                }

                // Insert an ending paragraph mark to give room for the wavy
                // red lines to do its work.
                wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);

                // Assign to the in/out range the end of the working range.
                setOutgoingRng(begRng.Start, wrkRng.End);
            }
            catch (Exception e)
            {
                Log.exception(e, "Failed in generate");
                throw new FTException("Failed in generate", e);
            }
        }
		public override bool preDisplay(Word.Range range)
		{
			if (!base.preDisplay(range)) 
			{
				return false;
			}

			IList tableViews = getWordTableView();

			foreach (TableView tv in tableViews) 
			{
				if (tv.getColumnCount() > 63) 
				{
					range.InsertAfter("The table being created has more columns than Microsoft Word allows (63).  Please break apart your table by inserting a manual page break.");
					setOutgoingRng(range);
					return false;
				}
			}

			return true;
		}

		public void CreationDate_TOC()
		{
			try
			{
				IEnumerator allTOC = tspdDoc_.getActiveWordDocument().TablesOfContents.GetEnumerator();

				while (allTOC.MoveNext())
				{
					Word.TableOfContents TOC_ = (Word.TableOfContents)allTOC.Current;
					//	TOC_ = tspdDoc_.getActiveWordDocument().TablesOfContents;
					TOC_.Update();

				}
				
				TspdTrial trial = tspdDoc_.getTspdTrial();
				FTDateTime dtCreated = trial.getCreateDate();
				string datecr =  dtCreated.getDateTime().ToString();
															
				FTDateTime.ClientDateFormat="yyyy MM dd";
				
				string sDate = dtCreated.getDateTime().ToShortDateString();

				string sCreated = dtCreated.getDateTime().ToString("dd MMMM yyyy");
				
				WordHelper.setVariableValue(tspdDoc_.getActiveWordDocument(),"tspd.trial.createdate",sCreated);


			}
			catch(Exception ex)
			{
				Log.exception(ex, "Error in updating table of contents/creation date.");
			}
			


		}
	}
}
