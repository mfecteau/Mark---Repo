using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

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
    internal sealed class TestTreatmentMacro
    {
        private static readonly string header_ = @"$Header: TestTreatmentMacro.cs, 1, 04-aug-10 12:05:10, Pinal Patel$";
    }
}

namespace TspdCfg.Purdue.DynTmplts
{
    /// <summary>
    /// Summary description for ContactDetailsMacro.
    /// </summary>
    public class TestTreatmentMacro : AbstractMacroImpl
    {
        public MacrosConfig m_MacrosConfig = null;

        public TestTreatmentMacro(MacroExecutor.MacroParameters mp) : base(mp)
        {
            //
            // TODO: Add constructor logic here
            //
        }

        #region Dynamic Tmplt Methods

        #region TestTreatmentMacro
        /// <summary>
        /// /// Displays contact information (Fax only) based on Role Type
        /// </summary>
        /// <param name="mp"></param>
        /// <returns></returns>
        public static MacroExecutor.MacroRetCd TestTreatment(MacroExecutor.MacroParameters mp)
        {
            #if false
            <ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.PregnancyMacro.Pregnancy,ProtocolDTs.dll" elementLabel="Pregnancy" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Population" autogenerates="true" toolTip="Pregnancy." shouldRun="true"/>
            #endif
            try
            {
                mp.pba_.setOperation("Test Treatment Macro", "Generating information...");

                TestTreatmentMacro macro = null;
                macro = new TestTreatmentMacro(mp);
                macro.preProcess();
                macro.displayTestTreatment();
                macro.postProcess();
                return macro.macroStatusCode_;
            }
            catch (Exception e)
            {
                Log.exception(e, "Error in Test Treatment Macro");
                mp.inoutRng_.Text = "Test Treatment Macro: " + e.Message;
            }
            return MacroExecutor.MacroRetCd.Failed;
        }

        #endregion

        #region RefTreatment
        /// <summary>
        /// Displays contact information based on Role Type
        /// </summary>
        /// <param name="mp"></param>
        /// <returns></returns>
        public static MacroExecutor.MacroRetCd RefTreatment(MacroExecutor.MacroParameters mp)
        {
            #if false
            <ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.PregnancyMacro.Pregnancy,ProtocolDTs.dll" elementLabel="Pregnancy" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Population" autogenerates="true" toolTip="Pregnancy." shouldRun="true"/>
            #endif
            try
            {
                mp.pba_.setOperation("Reference Treatment Macro", "Generating information...");

                TestTreatmentMacro macro = null;
                macro = new TestTreatmentMacro(mp);
                macro.preProcess();
                macro.displayTestTreatment();
                macro.postProcess();
                return macro.macroStatusCode_;
            }
            catch (Exception e)
            {
                Log.exception(e, "Error in Reference Treatment Macro");
                mp.inoutRng_.Text = "Reference Treatment Macro: " + e.Message;
            }
            return MacroExecutor.MacroRetCd.Failed;
        }

        #endregion

        public static MacroExecutor.MacroRetCd OtherTreatment(MacroExecutor.MacroParameters mp)
        {
            #if false
            <ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.PregnancyMacro.Pregnancy,ProtocolDTs.dll" elementLabel="Pregnancy" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Population" autogenerates="true" toolTip="Pregnancy." shouldRun="true"/>
            #endif
            try
            {
                mp.pba_.setOperation("Other Treatment Macro", "Generating information...");

                TestTreatmentMacro macro = null;
                macro = new TestTreatmentMacro(mp);
                macro.preProcess();
                macro.displayTestTreatment();
                macro.postProcess();
                return macro.macroStatusCode_;
            }
            catch (Exception e)
            {
                Log.exception(e, "Error in Other Treatment Macro");
                mp.inoutRng_.Text = "Other Treatment Macro: " + e.Message;
            }
            return MacroExecutor.MacroRetCd.Failed;
        }

        #endregion

        public void displayTestTreatment()
        {

            string chooserElementPath = this.macroEntry_.getElementPath();
            string fPath = tspdDoc_.getTrialProject().getTemplateDirPath() + "\\dyntmplts\\MacrosConfig.xml";
            m_MacrosConfig = new MacrosConfig(fPath, chooserElementPath);
            string msg = "";

            Word.Range inoutRange = this.startAtBeginningOfParagraph();
            Word.Range wrkRng = inoutRange.Duplicate;

            pba_.updateProgress(1.0);


            string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);
            BusinessObjectMgr bom_ = tspdDoc_.getBom();

            //Getting all the RoleTypes to select Test Articles.
            string[] roletypes = null;
            string strRoleType = m_MacrosConfig.getMessageByName("roletypes").Text;
            if (strRoleType.Contains("|"))
            {
                roletypes = strRoleType.Split('|');
            }
            else
            {
                roletypes = new string[] { strRoleType };
                //roletypes.SetValue(strRoleType, 0);
            }

            //Assuming, it will always be specifed in config file.
            ArrayList arr_Roletypes = new ArrayList();
            foreach (String rtype in roletypes)
            {
                if (rtype.Trim().Length > 0)
                {
                    arr_Roletypes.Add(rtype.Trim().ToLower());
                }
            } 

            List<Treatment> treatments = new List<Treatment>();
            foreach (Treatment treatment in this.bom_.getTreatments().Enumerable.OfType<Treatment>().OrderBy(tr => tr.getSequence()))
            {
                foreach (Component component in bom_.getAssociatedComponents(treatment).Enumerable)
                {
                    TestArticle testArticle = bom_.getTestArticle(component.AssociatedTestArticleID);
                    if ((testArticle != null) && arr_Roletypes.Contains(testArticle.PrimaryRole.ToLower()))
                        treatments.Add(treatment);
                }
                pba_.updateProgress(1.0);
            }

            if (treatments.Count <= 0)
            {
                msg = m_MacrosConfig.getMessageByName("exception1").Text;
                wrkRng.InsertAfter(msg);
                wrkRng.InsertParagraphAfter();
                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                // Set outgoing range
                inoutRange.End = wrkRng.End;
                setOutgoingRng(inoutRange);
                wdDoc_.UndoClear();
                return;
            }

            msg = m_MacrosConfig.getMessageByName("firstline").Text;
            if (msg.Trim().Length > 0)
            {
                wrkRng.InsertAfter(msg);
                wrkRng.InsertParagraphAfter();
                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
            }
            //Create table.
            wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

            Word.Table tbl = createTable(wrkRng, 6, treatments.Count + 1);
            //Applying Table Text Style.
            string oStyle = "";
            try
            {
                oStyle = m_MacrosConfig.getMessageByName("tabletext").Text;
                // object oStyle = (object)colStyle;
                tspdDoc_.getStyleHelper().setNamedStyle(oStyle, tbl.Range);
            }
            catch (Exception e)
            {
                Log.exception(e, e.Message + " Table Text -" + oStyle);
            }


            //First Column

            tbl.Cell(1, 1).Range.InsertAfter("Name");
            tbl.Cell(2, 1).Range.InsertAfter("Dosage Form");
            tbl.Cell(3, 1).Range.InsertAfter("Dosage Regimen");
            tbl.Cell(4, 1).Range.InsertAfter("Route");
            tbl.Cell(5, 1).Range.InsertAfter("Strength");
            tbl.Cell(6, 1).Range.InsertAfter("Supplier");

            //Start Writing with Second Columns
            string strTemp = "";

            int i = 2;
            foreach (Treatment treatment in this.bom_.getTreatments().Enumerable.OfType<Treatment>().OrderBy(tr => tr.getSequence()))
            {
                foreach (Component component in bom_.getAssociatedComponents(treatment).Enumerable)
                {
                    TestArticle testArticle = bom_.getTestArticle(component.AssociatedTestArticleID);
                    if ((testArticle == null) || !arr_Roletypes.Contains(testArticle.PrimaryRole.ToLower()))
                        continue;


                    TableCellWordFormat(wrkRng, tbl.Cell(1, i).Range, treatment.Name);
                    
                    strTemp = component.Formulation;
                    if (strTemp == "other")
                    {
                        strTemp = component.OtherFormulation;
                    }
                  
                    TableCellWordFormat(wrkRng, tbl.Cell(2, i).Range, strTemp);
                    TableCellWordFormat(wrkRng, tbl.Cell(3, i).Range, treatment.FrequencyUnit);

                    strTemp = treatment.DeliveryMethod;
                    if (strTemp == "other")
                    {
                        strTemp = treatment.OtherDeliveryMethod;
                    }

                    TableCellWordFormat(wrkRng, tbl.Cell(4, i).Range, strTemp);
                    TableCellWordFormat(wrkRng, tbl.Cell(5, i).Range, component.Strength);
                    TableCellWordFormat(wrkRng, tbl.Cell(6, i).Range, component.Manufacturer);

                    i++;  //Move to next column
                }
            }

            //Applying  Column Header Style, as Col Header = Row 1 

            try
            {
                oStyle = m_MacrosConfig.getMessageByName("colheader").Text;
                // object oStyle = (object)colStyle;
                tspdDoc_.getStyleHelper().setNamedStyle(oStyle, tbl.Rows[1].Range);
            }
            catch (Exception e)
            {
                Log.exception(e, e.Message + " Coloumn Header -" + oStyle);
            }

            //Applying  Row Header Style, Cell 1 of each Row.

            try
            {
                oStyle = m_MacrosConfig.getMessageByName("rowheader").Text;
                // object oStyle = (object)colStyle;
                for (int j = 1; j <= tbl.Rows.Count; j++)
                {
                    tspdDoc_.getStyleHelper().setNamedStyle(oStyle, tbl.Cell(j, 1).Range);
                }
            }
            catch (Exception e)
            {
                Log.exception(e, e.Message + " Row Header -" + oStyle);
            }

            wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

            // Set outgoing range
            inoutRange.End = wrkRng.End;
            setOutgoingRng(inoutRange);
            wdDoc_.UndoClear();
        }

        private static void TableCellWordFormat(Word.Range wrkRng, Word.Range table, string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                WordFormatter.FTToWordFormat2(ref wrkRng, text);
                wrkRng.Copy();
                table.Paste();
                wrkRng.Delete();
            }
        }
        
        public virtual Word.Table createTable(Word.Range viewRng, int rows, int cols)
        {

            // Turn off auto caption for Word tables.
            Word.AutoCaption ac = wdApp_.AutoCaptions.get_Item(ref WordHelper.AUTO_CAPTION_WORD_TABLE);
            bool oldState = ac.AutoInsert;
            ac.AutoInsert = false;

            Word.Range wrkRng = viewRng.Duplicate;


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


            tbl.Borders.Enable = VBAHelper.iTRUE;
            tbl.Borders.InsideLineWidth = MacroBaseUtilities.LINE_WIDTH_NORMAL;

            tbl.Borders.OutsideLineWidth = MacroBaseUtilities.LINE_WIDTH_NORMAL;
            // Reinstate auto caption for Word tables.
            ac.AutoInsert = oldState;

            tbl.Columns[1].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
            tbl.Columns[1].PreferredWidth = tbl.Application.InchesToPoints(0.84f);  //Fixed width

            // Increase viewRng to include the table.
            viewRng.End = tbl.Range.End;

            viewRng.Collapse(ref WordHelper.COLLAPSE_END);

            return tbl;
        }



        public override void postProcess()
        {
            // Clean up memory
        }
    }
}
