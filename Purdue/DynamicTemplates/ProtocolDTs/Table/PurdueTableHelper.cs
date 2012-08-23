using System;
using System.Collections;
using System.Reflection;
using Tspd.Businessobject;
using Tspd.MacroBase;
using Tspd.MacroBase.BaseImpl;
using Tspd.MacroBase.Table;
using Tspd.Icp;
using Tspd.Bridge;
using Tspd.Utilities;
using Tspd.Context;
using Tspd.Tspddoc;
using Word = Microsoft.Office.Interop.Word;

namespace TspdCfg.Purdue.DynTmplts.Table
{
	/// <summary>
	/// Summary description for PfizerTableHelper.
	/// </summary>
	public class PurdueTableHelper : TableHelper
	{
		// Collect the special events
		ArrayList pfizerContinuousEvents = null;
		ArrayList pfizerFowsToBeRemoved_ = new ArrayList();
		ArrayList pfizerFootNotesToAdd_ = new ArrayList(); 

		public PurdueTableHelper(AbstractTableDisplayMgr targetTableDispMgr) : base (targetTableDispMgr)
		{
		}

		public override void wordTableCreated(Word.Range wrkRng, Word.Table tbl, TableView tableView)
		{
			// Put here coz adorn table nukes this setting
		    tbl.Borders.OutsideLineWidth = MacroBaseUtilities.LINE_WIDTH_NORMAL;
			//tbl.Rows.LeftIndent = tbl.Application.InchesToPoints(0f);

            try
            {
                PurdueSOATableView rtv = tableView as PurdueSOATableView;
                TspdDocument currDoc = targetTableDispMgr_.TspdDoc;
                Hashtable dispSettings = new Hashtable();
                ArrayList headerRowSettings = new ArrayList();

                try
                {
                    //Setting Table Style as provided by clients.
                    currDoc.getStyleHelper().setNamedStyle(PurdueUtil.PURDUE_STYLE_TABLETEXT_10, tbl.Range);
                }
                catch (Exception ex)
                {
                    Log.exception(ex, ex.Message);
                }

                bool flag = false;

                string filepath = currDoc.getTrialProject().getTrialDirPath() + "\\" + currDoc.getDocumentDetails().getRelativeFileName();
                filepath = currDoc.getTrialProject().getTemplateDirPath() + "\\dyntmplts\\SOAConfig.txt";

                if (System.IO.File.Exists(filepath) == false)  //If Configuration file is present.
                {
                    filepath = BridgeProxy.getInstance().getSystemTemplatePath() + "\\dyntmplts\\SOAConfig.txt"; 
                }

                string currLine = "";
                string header = "", val = "";

                if (System.IO.File.Exists(filepath) == true)  //If Configuration file is present.
                {
                    //System.Windows.Forms.MessageBox.Show(filepath);
                    System.IO.StreamReader strReader = new System.IO.StreamReader(filepath);
                    //strReader.Peek(
                    while (strReader.Peek() >= 0)
                    {
                        currLine = strReader.ReadLine();
                        // compare currLine to Report Type and set FLAG = TRUE for reading lines then after.
                        if (currLine.IndexOf("TID") >= 0)
                        {
                            currLine = currLine.Substring(currLine.IndexOf("=") + 1);
                            if (currLine != "")
                            {
                                if (currDoc.getTspdTrial().getTspdTemplateTid().ToString() == currLine.Trim())
                                {
                                    flag = true;
                                }
                            }
                        }
                        else  //IF not first line.
                        {
                            if ((flag == true) && (currLine.Trim().Length > 0))
                            {

                                if (currLine.Trim().ToUpper() == "ETID")
                                {
                                    break;
                                }

                                header = currLine.Substring(0, currLine.IndexOf("="));
                                val = currLine.Substring(currLine.IndexOf("=") + 1);
                                //System.Windows.Forms.MessageBox.Show(header);
                                if (header.ToUpper() != "ETID")
                                {
                                    //dispSettings.Add(header.Trim(),val.Trim());
                                    headerRowSettings.Add(val.Trim());
                                }

                            }
                        }
                    }
                }
                else
                {
                    headerRowSettings.Add("true");
                    headerRowSettings.Add("true");
                    headerRowSettings.Add("true");
                    headerRowSettings.Add("true");
                    headerRowSettings.Add("true");
                    headerRowSettings.Add("true");

                }

                int eventrow = 1;

                if (headerRowSettings[0].ToString().ToUpper() == "FALSE")
                {//Period Row
                    tbl.Rows[eventrow].Delete();                   
                }
                else
                {
                    eventrow++;
                }


                if (headerRowSettings[1].ToString().ToUpper() == "FALSE")
                {//Period Row
                    tbl.Rows[eventrow].Delete();
                }
                else
                {
                    eventrow++;
                }

                if (rtv.HasSubPeriods)
                {
                    if (headerRowSettings[2].ToString().ToUpper() == "FALSE")
                    {//Sub-Period Row

                        tbl.Rows[eventrow].Delete();
                    }
                    else
                    {
                        eventrow++;
                    }
                }


                if (headerRowSettings[3].ToString().ToUpper() == "FALSE")
                {//Visit Row

                    tbl.Rows[eventrow].Delete();
                }
                else
                {
                    eventrow++;
                }

                if (rtv.HasStudyDays)
                {
                    if (headerRowSettings[4].ToString().ToUpper() == "FALSE")
                    {//Study Day Row

                        tbl.Rows[eventrow].Delete();
                    }
                    else
                    {
                        eventrow++;
                    }
                }

                if (rtv.HasVisitWindow)
                {
                    if (headerRowSettings[5].ToString().ToUpper() == "FALSE")
                    {//Window Row

                        tbl.Rows[eventrow].Delete();
                    }
                    else
                    {
                        eventrow++;
                    }
                }

                // Remove cell border for apearance of cell merge
                int startRow = 0;
                int endRow = 0;

                if (headerRowSettings[0].ToString().ToUpper() != "TRUE")
                {
                    startRow = 1;
                    endRow = 1;
                }
                else
                {
                    startRow = 2;
                    endRow = 2;
                }


                if (headerRowSettings[1].ToString().ToUpper() == "TRUE") endRow++;  //Period
                if ((rtv.HasSubPeriods) && (headerRowSettings[2].ToString().ToUpper() == "TRUE")) endRow++;		 //Sub Period

                if (headerRowSettings[3].ToString().ToUpper() == "TRUE") endRow++;  //Visitrow

                //				if (rtv.HasStudyDays) endRow++;
                //				if (rtv.HasVisitWindow) endRow++;

                for (int i = startRow; i < endRow; i++)
                {
                    // Hide the borders to simulate a cell merge
                    Word.Cell c1 = tbl.Rows[i].Cells[1];

                    c1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    if ((i + 1) < endRow)
                    {
                        Word.Cell c2 = tbl.Rows[i + 1].Cells[1];
                        c1.Borders[Word.WdBorderType.wdBorderBottom].Visible = true;
                        c2.Borders[Word.WdBorderType.wdBorderTop].Visible = false;
                    }
                }

                // This will break coz code later access tlb.rows will break if there are merged things
                // c1.Merge(c2);

                if ((rtv.HasStudyDays) && (headerRowSettings[4].ToString().ToUpper() == "TRUE"))
                {
                    //endRow++;
                    Word.Cell c1 = tbl.Rows[endRow].Cells[1];
                    c1.Range.Italic = -1;
                    endRow++;
                    //c1.Range.Bold =-1;		
                }
                //Code for putting double line border for last Header row
               

                if ((rtv.HasVisitWindow) && (headerRowSettings[5].ToString().ToUpper() == "TRUE"))
                {
                    if (headerRowSettings[0].ToString().ToUpper() == "FALSE")
                    {
                        endRow++;
                    }
                }

                if (headerRowSettings[0].ToString().ToUpper() == "FALSE")
                {
                    endRow--;
                }
                eventrow--;


                Word.Row rEnd = tbl.Rows[eventrow];
               
                rEnd.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleDouble;
                rEnd.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth075pt;


                //for applying BOLD to HEADER ROWS -  PINAL PATEL - 4/18/2010
                try
                {
                    object rStyle;
                    for (int i = 1; i <= eventrow; i++)
                    {
                        Word.Row currHeaderRow = tbl.Rows[i];
                        if (i <= eventrow)
                        {
                           currHeaderRow.Range.Font.Bold = -1;  //setting it to Bold
                        }
                    }
                }
                catch (Exception ex)
                {
                    //Bypass error, if style is missing in document.
                    //System.Windows.Forms.MessageBox.Show(ex.ToString());
                    Log.exception(ex, "Error in setting styles for header rows.");
                }

                // Left align first column, center the rest
                int ithRow;

                for (ithRow = 1; ithRow <= tbl.Rows.Count; ++ithRow)
                {
                    Word.Range rng = tbl.Rows[ithRow].Range;
                    //rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    //rng.Move(ref WordHelper.CELL, ref MacroBaseUtilities.O1);
                    //rng.MoveEnd(ref WordHelper.ROW, ref MacroBaseUtilities.O1);
                    rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

                SOA _soa = tableView.DataModel.getSOA();
                foreach (Word.Row rw in tbl.Rows)
                {
                    try
                    {
                        if (rw.Cells.Count > 1)
                        {
                            Word.Cell cl = rw.Cells[1];
                            cl.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            string taskname = cl.Range.Text.Trim();
                            if (taskname.Length > 0)
                            {
                                Task tsk = PurdueUtil.getTaskByName(_soa, taskname);
                                if (tsk != null)
                                {
                                    TaskDivider tdParent = _soa.getParentTaskDivider(tsk);
                                    if (tdParent != null)
                                    {
                                        Word.Range cr = cl.Range.Duplicate;
                                        cr.Collapse(ref WordHelper.COLLAPSE_START);
                                        cr.ParagraphFormat.LeftIndent = cr.Application.InchesToPoints(0.25f);
                                        cr.InsertAfter("");
                                    }
                                }
                            }
                        }
                        else
                        {   /// All the taskHeader rows need to be left alligned.
                            rw.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.exception(ex, "Error setting task with divider - TableHelper");
                    }
                }

                tbl.Cell(1, 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
			

            }
            catch (Exception ex)
            {
                Log.exception(ex, "Error Deleting/Merging");
            }


            IFootnoter footNoter = tableView.getFootnoter();

            if (footNoter.hasFootnotes())
            {
                Word.Table fnTbl = createFootnoteTable(wrkRng, footNoter.getFootnotes().Count);
                fillFootnoteTable(fnTbl, footNoter.getFootnotes(), tableView);
            }
		}

		private Word.Table createFootnoteTable(Word.Range viewRng, int numFootnotes) 
		{
			// Turn off auto caption for Word tables.
			Word.AutoCaption ac = targetTableDispMgr_.WordApp.AutoCaptions.get_Item(ref WordHelper.AUTO_CAPTION_WORD_TABLE);
			bool oldState = ac.AutoInsert;
			ac.AutoInsert = false;

			Word.Range wrkRng = viewRng.Duplicate;
			// Collapse the range to the end of the paragraph mark so that the table can be added
			// after it.
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

            
			bool pageBreak = false;
			if (pageBreak)
			{
				object wdPageBreak = Word.WdBreakType.wdPageBreak;
				wrkRng.InsertBreak(ref wdPageBreak);
			}
			else
			{
				wrkRng.InsertParagraphAfter();
			}

			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			int nbrCols = 2;
			int nbrRows = numFootnotes;

			// Insert the table. Note, that the table is inserted starting at but after the range.
			// So viewRng isn't increased.
			Word.Table tbl =
				targetTableDispMgr_.WordDoc.Tables.Add(
				wrkRng, nbrRows, nbrCols,
				ref WordHelper.WORD8_TABLE_BEHAVIOR, ref VBAHelper.OPT_MISSING);

			// Reinstate auto caption for Word tables.
			ac.AutoInsert = oldState;

			tbl.Borders.Enable = VBAHelper.iFALSE;
			tbl.LeftPadding = 0;
			tbl.RightPadding = 0;
			tbl.Rows.LeftIndent = tbl.Application.InchesToPoints(0f);

			tbl.Rows.AllowBreakAcrossPages = VBAHelper.iFALSE;

            try
            {
                object oStyle = targetTableDispMgr_.TspdDoc.getStyleHelper().setNamedStyle(PurdueUtil.PURDUE_STYLE_FOOTNOTE, tbl.Range);
                //wrkRng.set_Style(ref oStyle);
            }
            catch (Exception ex)
            { }


			// Make first row have a good size
			foreach (Word.Row rw in tbl.Rows)
			{
				Word.Cell cl = rw.Cells[1];
				cl.SetWidth(tbl.Application.InchesToPoints(0.25f), 
					Word.WdRulerStyle.wdAdjustProportional);
			}

			// Increase viewRng to include the table.
			viewRng.End = tbl.Range.End;

			targetTableDispMgr_.WordDoc.UndoClear();

			return tbl;
		}

		private class FNComparer : IComparer
		{
			public int Compare(object x, object y)
			{
				FootNoteWrapper fn1 = x as FootNoteWrapper;
				FootNoteWrapper fn2 = y as FootNoteWrapper;

				return fn1.footNoteNumber.CompareTo(fn2.footNoteNumber);
			}
		}

		private void fillFootnoteTable(Word.Table tbl, Hashtable footNotes, TableView tableView) 
		{
			Word.Range wrk;
			FNComparer fnComparer = new FNComparer();

			int i = 0;
			ArrayList al = new ArrayList(footNotes.Values);
			al.Sort(fnComparer);

			foreach (FootNoteWrapper fnw in al)
			{
				Word.Row row = tbl.Rows[i + 1];
				// Number column
				wrk = row.Cells[1].Range;
				wrk.Paragraphs.KeepTogether = VBAHelper.iTRUE;
				wrk.Paragraphs.KeepWithNext = VBAHelper.iTRUE;
				wrk.End--;
				
				string s = fnw.footNoteNumberString;
				wrk.Text = s;

				// Footnote text
				wrk = row.Cells[2].Range;
				wrk.Paragraphs.KeepTogether = VBAHelper.iTRUE;
				wrk.Paragraphs.KeepWithNext = VBAHelper.iTRUE;
				wrk.End--;

				//wrk.Text =  Formatter.stripFormatInstruction(fnw.footNote.getFootNoteText());
				WordFormatter.FTToWordFormat2(ref wrk, fnw.footNote.getFootNoteText());

				targetTableDispMgr_.WordDoc.UndoClear();

				i++;
			}

			

			SOATableFormat tblFmt = targetTableDispMgr_.TspdDoc.getSOATblFormat(tableView.DataModel.getSOA().getObjID());

			Word.Font targetFont;
			if (tblFmt == null)
			{
				// create a fake one to run thru defaults...
				TrialDocument.SOATableFormatCV cv = new TrialDocument.SOATableFormatCV();
				tblFmt = cv.newSOATableFormat();
			}

			targetFont = tbl.Range.Font;

			if (tblFmt.getDocTableBodyFontName() != null)
			{
				targetFont.Name = tblFmt.getDocTableBodyFontName();
			}

			if (tblFmt.getDocTableBodyFontSize() != null)
			{
				targetFont.Size = tblFmt.getDocTableBodyFontSize();
			}

//			targetFont.Bold = (tblFmt.getDocTableBodyFontBold() ? -1 : 0);	
//			targetFont.Italic = (tblFmt.getDocTableBodyFontItalics() ? -1 : 0);	

			targetTableDispMgr_.WordDoc.UndoClear();
		}

		public override Word.Table createTable(Word.Range viewRng, TableView tableView) 
		{
			DefSOADataModel dataModel = tableView.DataModel;
			SOA soa = dataModel.getSOA();

			// Collect the special events
			pfizerContinuousEvents = soa.getStripedEvents();

			return base.createTable(viewRng, tableView);
		}

		private ArrayList pfizerGetRowsToBeRemovedList()
		{
			return this.pfizerFowsToBeRemoved_;
		}

		private ArrayList pfizerGetAddFootNoteActionsList()
		{
			return this.pfizerFootNotesToAdd_;
		}

		public override void applyFootNotes(Word.Table tbl, TableView tableView)
		{
			//set Ref style to Footnote style
			object fnRef = "Footnote Reference";
			object fnText = "Footnote Text";
			Word.Font  fnTextFont = targetTableDispMgr_.WordDoc.Styles.get_Item(ref fnText).Font;
			//copy attributes
			targetTableDispMgr_.WordDoc.Styles.get_Item(ref fnRef).Font.Size = fnTextFont.Size;
			 
			pfizerGetAddFootNoteActionsList().Sort(); //sort it based on the IComparable of AddFootnoteAction

			IEnumerator fnEnum = pfizerGetAddFootNoteActionsList().GetEnumerator();
			while (fnEnum.MoveNext())
			{
				AddFootNoteAction action = fnEnum.Current as AddFootNoteAction;
				if (action != null)
				{
					//System.Diagnostics.Debug.WriteLine(action.ToString());
					action.applyFootnote(); //apply the footnote
				}
			}
		}

		public override void purgeExcludedRows(Word.Table table, TableView tableView)
		{
			// now remove the unwanted rows, moving backwards
			pfizerGetRowsToBeRemovedList().Sort();
			for (int i= pfizerGetRowsToBeRemovedList().Count - 1; i >= 0 ;i--)
			{
				int row = (int)pfizerGetRowsToBeRemovedList()[i];
				try
				{
					table.Rows[row].Delete();
				}
				catch(Exception ex)
				{
					Log.exception(ex, "Error deleting Word Table Row: " + row);
				}
			}
			targetTableDispMgr_.WordDoc.UndoClear();
		}

		public override void fillTableRowHeaders(Word.Table table, TableView tableView, IComparer comparer)
		{
			// merge cells from bottom up...
			int headerRowLength = tableView.getHeaderRowCount();
			for (int row = headerRowLength -1; row >= 0; row--)
			{
				//check to see if the TableView want's the rule processed first...
				if (!tableView.processAndDisplayRow(row))
				{
					pfizerGetRowsToBeRemovedList().Add(row + 1);
					continue; 
				}

				// ignore the the Column Header (index = 0)
				int colLength = tableView.getColumnCount();
				ArrayList mergeRangesList = new ArrayList();
				int begIndex = 0;
				object begObj = null;
				object curObj = null;
				for (int col = 1; col < colLength; col++)
				{
					curObj = tableView.getValueAt(row,col);
					if (begObj == null)
					{
						begIndex = col;
						begObj = curObj;
					}

					if (begObj != null &&
						/*use a comparer if one was passed in, else use Equals*/
						(!MacroBaseUtilities.checkEquality(begObj, curObj, comparer) ||
						/*last one*/
						col == (colLength -1)) )
					{
						IndexRange newRange = new IndexRange();
						newRange.BegRange = begIndex;
						if (col == (colLength -1))
						{
							if (MacroBaseUtilities.checkEquality(begObj, curObj, comparer))
							{//last cell needs to be merged too
								newRange.EndRange = col;
								mergeRangesList.Add(newRange);
								//done...
							}
							else
							{//last cell needs to be merged in a separate range so that it can be "displayed" below
								newRange.EndRange = col - 1;
								mergeRangesList.Add(newRange);
								//add an extra one for the last cell
								newRange = new IndexRange();
								newRange.BegRange = col; //last col
								newRange.EndRange = col; //ditto
								mergeRangesList.Add(newRange);
								//done...
							}
						}
						else
						{
							newRange.EndRange = col - 1;
							mergeRangesList.Add(newRange);
							//set next range parm
							begIndex = col;
							begObj = curObj;
						}
					}
				}//end for...

				// now merge the collection found in reverse order, because
				// the cell indexes will be whacked
				for (int k = (mergeRangesList.Count - 1); k >= 0; k--)
				{
					IndexRange aRange = (IndexRange)mergeRangesList[k];
					Word.Cell begCell = table.Cell(row + 1, aRange.BegRange + 1);
					
					if (aRange.BegRange < aRange.EndRange)
					{	// only merge when they are different
						Word.Cell endCell = table.Cell(row + 1, aRange.EndRange + 1);
						begCell.Merge(endCell);
					}
					// now apply renderer on that merged cell
					object data = tableView.getValueAt(row,aRange.BegRange);
					RowHeaderCell headerCell = tableView.getRowHeaderCell(row, data.GetType());
					if (headerCell != null)
					{
						//only if one is registered
						//apply table cell data to the row
						headerCell.display(targetTableDispMgr_.TspdDoc, begCell, data);
						//headerCell.addFootNotes(tableView.getFootnoter(),begCell,data);
						pfizerGetAddFootNoteActionsList().Add(new AddFootNoteAction(begCell,tableView.getFootnoter(),data,headerCell));
					}
					
				}

			}//end for...
			
			targetTableDispMgr_.WordDoc.UndoClear();
		}

		public override void fillTableDetail(Word.Table tbl, TableView tableView) 
		{
			
			pfizerGetRowsToBeRemovedList().Clear(); //clear before filling up the tableDetails/header
			pfizerGetAddFootNoteActionsList().Clear(); //clear the Add footnotes collection

			Word.Row curRow;
			ColumnHeaderCell colCell;
			TableCell dataCell;
			Word.Cell cell;
			object data;
			Word.Rows tableRows = tbl.Rows;

			//don't process the header rows...
			int totalRowCount = tableView.getHeaderRowCount() +  tableView.getDataRowCount();
			for (int i = tableView.getHeaderRowCount(); i < totalRowCount; i++)
			{
				// check to see if the TableView want's the rule processed first...
				if (!tableView.processAndDisplayRow(i))
				{
					pfizerGetRowsToBeRemovedList().Add(i + 1);
					continue; 
				}

				curRow = tableRows[i + 1];		
				for (int j = 0; j < tableView.getColumnCount(); j++)
				{
					data = tableView.getValueAt(i,j);
					if (data == null)
					{
						continue; //no data.. keep going
					}

					cell = curRow.Cells[j + 1];					
					if (j == 0)
					{
						// column 0, header column
						colCell = tableView.getColumnHeaderCell(data.GetType());
						// apply header col to cell
						if (colCell != null)
						{
							// only if one is registered
							colCell.display(targetTableDispMgr_.TspdDoc, cell,data);
							pfizerGetAddFootNoteActionsList().Add(new AddFootNoteAction(cell,tableView.getFootnoter(),data,colCell));
						}
					}
					else
					{
						dataCell = tableView.getTableCell(data.GetType());
						if (dataCell != null)
						{
							// only if one is registered
							// apply table cell data to the row
							dataCell.display(targetTableDispMgr_.TspdDoc,cell,data);
							pfizerGetAddFootNoteActionsList().Add(new AddFootNoteAction(cell,tableView.getFootnoter(),data, dataCell));
						}
					}
				}

				tableView.rowCreated(curRow); //callback to TableView if it wan't to customize row...
			}

			targetTableDispMgr_.WordDoc.UndoClear();

			pfizerDisplayContinuousEvents(tbl, tableView);

			
		}

		private void pfizerDisplayContinuousEvents(Word.Table tbl, TableView tableView) 
		{
			DefSOADataModel dataModel = tableView.DataModel;
			SOA soa = dataModel.getSOA();
			CEvtTableCell cevtCell = new CEvtTableCell();

			ScheduleTree soaTreeRoot = ScheduleTree.buildScheduleTree(soa);
			BusinessObjectMgr theBom = ContextManager.getInstance().getActiveDocument().getBom();
			
			// soaTreeRoot.dumpTree(theBom);

			/*
			// outs
			object[,] matrix = null;
			int columns = 0;
			int rows = 0;

			soaTreeRoot.populateMatrix(ScheduleTree.TIME_CHUNK, out matrix, out rows, out columns);
			ScheduleTree.dumpMatrix(matrix, rows, columns, theBom);
			*/
			
			int cevtRow = tableView.getHeaderRowCount() + tableView.getDataRowCount() + 1;
			int firstRow = cevtRow;

			// Process them
			foreach (ProtocolEvent cevt in pfizerContinuousEvents) 
			{
				Word.Row row = tbl.Rows[cevtRow];
				Word.Cell cell = row.Cells[1];

				MacroBaseUtilities.putElemRefInCell(
					targetTableDispMgr_.TspdDoc, cell, cevt, 
					ProtocolEvent.DISPLAYNAME, true);

				pfizerGetAddFootNoteActionsList().Add(new AddFootNoteAction(
					cell, tableView.getFootnoter(), cevt, cevtCell));

				if (cevtRow == firstRow) 
				{
					row.Borders[Word.WdBorderType.wdBorderTop].LineWidth =
						MacroBaseUtilities.LINE_WIDTH_HEAVYNORMAL;
				}

				cevtRow++;

				int startCellCol = -1;
				int endCellCol = -1;

				Word.Cell startCell = null;
				Word.Cell endCell = null;

				Activity actStart = null;
				Activity actEnd = null;

				// Find the start and end anchors for the continuous event
				soa.getStartEndActivities(cevt, ref actStart, ref actEnd);

				for (int i = 1; i < tableView.getColumnCount(); i++) 
				{
					object visit = dataModel.getRowHeaderData()[DefSOADataModel.VISIT_ROW, i];
					if (visit is ProtocolEvent) 
					{
						ProtocolEvent pe = visit as ProtocolEvent;

						if (actStart != null && startCell == null &&
							(actStart.getObjID() == pe.getObjID() ||
							soaTreeRoot.isChild(actStart.getObjID(), pe.getObjID())))
						{
							startCell = row.Cells[i + 1];
							startCellCol = i + 1;
						}

						if (actEnd != null &&
							(actEnd.getObjID() == pe.getObjID() ||
							soaTreeRoot.isChild(actEnd.getObjID(), pe.getObjID())))
						{
							endCell = row.Cells[i + 1];
							endCellCol = i + 1;
						}
					}
				}

				string beginArrow = "\u25C4";
				string endArrow = "\u25BA";
				string dash = "-";

				string startText = beginArrow + dash;
				string endText = dash + endArrow;

				// Merge these cells
				if (startCell == null) 
				{
					startText = dash + dash;
					startCell = row.Cells[2];
					startCellCol = 2;
				}

				if (endCell == null) 
				{
					endText = dash + dash;
					endCell = row.Cells[tableView.getColumnCount()];
					endCellCol = tableView.getColumnCount();
				}

				if (startCellCol != endCellCol) 
				{
					startCell.Merge(endCell);
				}

				startCell.Range.Text = startText + endText;
				startCell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
			}

			targetTableDispMgr_.WordDoc.UndoClear();
		}
	}
}
