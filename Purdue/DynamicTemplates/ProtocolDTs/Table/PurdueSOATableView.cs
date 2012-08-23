using System;
using Tspd.MacroBase;
using Tspd.MacroBase.Table;
using Tspd.MacroBase.BaseImpl;
using Tspd.Macros;
using Tspd.Utilities;
using Tspd.Businessobject;
using Word = Microsoft.Office.Interop.Word;

namespace TspdCfg.Purdue.DynTmplts.Table
{
	/// <summary>
	/// Summary description for PfizerSOATableView.
	/// </summary>
	public class PurdueSOATableView : DefSOATableView
	{
		private int lastHeaderRow_ = 1;
		bool _isBlinded = false;

		public PurdueSOATableView()
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

		public int LastHeaderRow
		{
			get { return lastHeaderRow_; }
		}


		public bool HasSubPeriods
		{
			get { return hasSubPeriods_; }
		}

		public bool HasStudyDays
		{
			get { return hasStudyDays_; }
		}

		public bool HasVisitWindow
		{
			get { return hasVisitWindow_; }
		}

		public override void registerCellRenderers()
		{
			base.registerCellRenderers();

			// Register by class for row header
			cellRendererLookup_.registerRowHeaderCell(DefSOADataModel.PERIOD_ROW, typeof(Period), new PfizerPeriodRowCellRenderer());
			cellRendererLookup_.registerRowHeaderCell(DefSOADataModel.SUBPERIOD_ROW, typeof(SubPeriodWrapper), new PfizerSubPeriodRowCellRenderer());
			cellRendererLookup_.registerRowHeaderCell(DefSOADataModel.STUDYDAYS_ROW, typeof(VisitStudyDays), new PfizerSOAVisitStudyDaysRowHeaderCell());
			cellRendererLookup_.registerRowHeaderCell(DefSOADataModel.VISIT_ROW, typeof(ProtocolEvent), new PfizerVisitRowCellRenderer());
			cellRendererLookup_.registerRowHeaderCell(DefSOADataModel.VISITWINDOW_ROW, typeof(VisitWindow), new PfizerVisitWindowRowHeaderCell());

			// Register for row number for row header
			cellRendererLookup_.registerRowHeaderCell(DefSOADataModel.PERIOD_ROW, null, new PfizerPeriodRowCellRenderer());
			cellRendererLookup_.registerRowHeaderCell(DefSOADataModel.SUBPERIOD_ROW, null, new PfizerSubPeriodRowCellRenderer());
			cellRendererLookup_.registerRowHeaderCell(DefSOADataModel.STUDYDAYS_ROW, null, new PfizerSOAVisitStudyDaysRowHeaderCell());
			cellRendererLookup_.registerRowHeaderCell(DefSOADataModel.VISIT_ROW, null, new PfizerVisitRowCellRenderer());
			cellRendererLookup_.registerRowHeaderCell(DefSOADataModel.VISITWINDOW_ROW, null, new PfizerVisitWindowRowHeaderCell());

			// Register by class for non-header
			cellRendererLookup_.registerTableCell(typeof(TaskVisitNot), new NoValueRenderer(Word.WdColor.wdColorWhite));

			cellRendererLookup_.registerColumnHeaderCell(typeof(TaskWrapper), new PfizerTaskColumnHeaderCell(this));
		}

		public override IFootnoter getFootnoter()
		{
			if (footNoter_ == null)
			{
				footNoter_ = new PurdueFootnoter(
					parentTableDisplayMgr_.BusObjMgr, parentTableDisplayMgr_.WordDoc);
			}

			return footNoter_;
		}

		public override void formatTable(Word.Table tbl)
		{
		//	object oStyle = ParentTableDisplayMgr.TspdDoc.getStyleHelper().setNamedStyle(PurdueUtil.PURDUE_STYLE_TABLETEXT_10, tbl.Range);


			int ithRow;
			// vertical center cells
			tbl.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

			// Border throught the table.
			tbl.Borders.Enable = VBAHelper.iTRUE;
			tbl.Borders.InsideLineWidth = MacroBaseUtilities.LINE_WIDTH_NORMAL;
			tbl.Borders.OutsideLineWidth = MacroBaseUtilities.LINE_WIDTH_NORMAL;

			tbl.Rows.AllowBreakAcrossPages = VBAHelper.iFALSE;

			// Left align first column, center the rest
			tbl.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
			for (ithRow = 1; ithRow <= tbl.Rows.Count; ++ithRow) 
			{
				Word.Range rng = tbl.Rows[ithRow].Cells[1].Range;
				rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
				rng.Move(ref WordHelper.CELL, ref MacroBaseUtilities.O1);
				rng.MoveEnd(ref WordHelper.ROW, ref MacroBaseUtilities.O1);
				rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
			}
			
			ParentTableDisplayMgr.WordDoc.UndoClear(); //sprinkle


			ParentTableDisplayMgr.WordDoc.UndoClear();

			lastHeaderRow_ = 1;
			for (int i= 0; i < getHeaderRowCount(); i++)
			{
				if (processAndDisplayRow(i))
					lastHeaderRow_ = i + 1;
			}

			if (ParentTableDisplayMgr.useHeaderRows()) 
			{
				// Specify the header rows. (Note, this may not work and I don't really know how
				// to test it other than creating more than 1 page full of tasks!)
				for (ithRow = 1; ithRow <= lastHeaderRow_; ++ithRow) 
				{
					tbl.Rows[ithRow].HeadingFormat = VBAHelper.iTRUE;
				}
			}
			
			ParentTableDisplayMgr.WordDoc.UndoClear();
		}

		public override bool processAndDisplayRow(int row)
		{
			if (row == DefSOADataModel.SUBPERIOD_ROW)
			{	
				return hasSubPeriods_;
			}

			if (row == DefSOADataModel.STUDYDAYS_ROW)
			{	
				return hasStudyDays_;
			}

			if (row == DefSOADataModel.VISITWINDOW_ROW)
			{
				return hasVisitWindow_;
			}

			int headerRowCount = getHeaderRowCount();
			int totalRowCount = headerRowCount + getDataRowCount();
			
			//test task rows to ensure they have TaskVisits
			if (row >= headerRowCount && row < totalRowCount)
			{
				TaskWrapper tw = getValueAt(row, DefSOADataModel.TASK_COLUMN) as TaskWrapper;
				
				if (tw != null && tw.getWrapped() is TaskDivider)
				{
					return true;  //always show the TaskDivider	
				}
				if(_isBlinded == true && tw != null && tw.getWrapped() is Task)
				{
					Task t = (Task)tw.getWrapped();
					if(t.isDosingTask())
					{
						return false;
					}
				}

				if (getLeadingPageBreak() != null && getLeadingPageBreak().getHideUnusedTasks() && 
					tw != null && !tw.rangeContainsTaskVisits(curResults_.LowerX,curResults_.UpperX))
				{
					return false;
				}

				return true;
			}
			
			return true;
		}

		public override string getTitleText()
		{
			string tableCaption = "";
			tableCaption = String.Format(TSPDKeys.DEFAULT_TABLE_VIEW.toLocalizedString(), getTableViewNumber().ToString());
			tableCaption += VBAHelper.TAB + this.dataModel_.getSOA().getActualDisplayValue();
			if (parentTableDisplayMgr_.hasMultipleLogicalTables())
			{
				tableCaption += ": " + this.dataModel_.getSOA().getActualDisplayValue();
			}
			if (this.dataModel_.getSOA().isSchemaDesignMode()) 
			{ 
				long armID = this.dataModel_.getSOA().getArmInFocus();
				Arm arm = ParentTableDisplayMgr.BusObjMgr.findArmByID(armID);
				tableCaption += ", for Arm: " + arm.getActualDisplayValue();
			}
			if (curFilter_.LowerPeriodBreak != null && !MacroBaseUtilities.isEmpty(curFilter_.LowerPeriodBreak.getPageName())) 
			{
				tableCaption += ", " + curFilter_.LowerPeriodBreak.getPageName().Trim();
				// append the pgBreak stuff too...
				if (curFilter_.LowerTaskBreak != null && !MacroBaseUtilities.isEmpty(curFilter_.LowerTaskBreak.getPageName()))
				{
					tableCaption += " / " + curFilter_.LowerTaskBreak.getPageName().Trim();
				}
			}
			else if (curFilter_.LowerTaskBreak != null && !MacroBaseUtilities.isEmpty(curFilter_.LowerTaskBreak.getPageName()))
			{
				tableCaption += ", " + curFilter_.LowerTaskBreak.getPageName().Trim();
			} 
			return tableCaption;
		}

		public override void adornTable(Word.Table table)
		{
			// Border around the table, all our row purging and processing
			//may have nuked it.
			table.Borders.OutsideLineWidth = MacroBaseUtilities.LINE_WIDTH_THICK;

			bool useMethod1 = false;

			if (ParentTableDisplayMgr.addCaptionRow()) 
			{
				//put table Caption on the first row
				object fRow = table.Rows[1];
				Word.Row tRow = table.Rows.Add(ref fRow);
				Word.Cell cell = tRow.Cells[1];
				Word.Font rFont = cell.Range.Font;
				rFont.Bold = -1; //bold it
				rFont.Italic = 0; //no italics
				//rFont.Size = 10; //size it to 10
				// tRow.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth025pt;
				// tRow.Borders.OutsideColor= Word.WdColor.wdColorWhite;
				tRow.Borders.InsideLineStyle=Word.WdLineStyle.wdLineStyleNone;
				


				cell.Range.Text = getTitleText();
				Word.Range rowRange = tRow.Range;
				tRow.Cells.Merge();
				rowRange.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

				useMethod1 = true;
			}

			if (useMethod1) 
			{
				// Method #1. Insert a bookmark with a well-known name around the title for DocDiff.
				Word.Row tRow = table.Rows[1];
				Word.Range tRng = tRow.Cells[1].Range.Duplicate;
				tRng.End--;
				Word.Bookmark tBm =
					DocDiffHelper.insertSOATableTitleBookmark(
					this.ParentTableDisplayMgr.TspdDoc.getActiveWordDocument(),
					tRng);
			}
			else
			{
				// Method #2. Insert a bookmark with a well-known name which includes the title.
				string soaTableTitle = getTitleText();
				if (soaTableTitle.Length > 28) 
				{
					soaTableTitle = soaTableTitle.Substring(0, 28);
				}

				Word.Bookmark tBm =
					DocDiffHelper.insertSOATableTitleBookmark(
					this.ParentTableDisplayMgr.TspdDoc.getActiveWordDocument(),
					table,
					soaTableTitle);
			}
		}

		public override void rowCreated(Word.Row row)
		{
			//check to see if it's a TaskDivider... if so merge the cells
			TaskWrapper tw = getValueAt(row.Index - 1, DefSOADataModel.TASK_COLUMN) as TaskWrapper;
			if (tw != null && tw.getWrapped() is TaskDivider)
			{
				try
				{
					Word.Cell begCell = row.Cells[1];
					Word.Cell endCell = row.Cells[row.Cells.Count];

					begCell.Merge(endCell);
					begCell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

					/*
					begCell.Shading.Texture = Word.WdTextureIndex.wdTextureDiagonalUp;
					begCell.Shading.ForegroundPatternColor = Word.WdColor.wdColorAutomatic;
					begCell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray50;
	
					row.Borders.Item(Word.WdBorderType.wdBorderBottom).LineWidth =
						MacroBaseUtilities.LINE_WIDTH_HEAVYNORMAL;
					row.Borders.Item(Word.WdBorderType.wdBorderTop).LineWidth =
						MacroBaseUtilities.LINE_WIDTH_HEAVYNORMAL;
					*/
				
				}
				catch(Exception ex)
				{
					//cell may not be there so move on...
				}
			}
		}
	}
}
