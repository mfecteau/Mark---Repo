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
    internal sealed class ContactDetailsMacro
	{
        private static readonly string header_ = @"$Header: ContactDetailsMacro.cs, 1, 20-jul-10 12:05:10, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
    /// Summary description for ContactDetailsMacro.
	/// </summary>
	public class ContactDetailsMacro : AbstractMacroImpl
	{
        public ContactDetailsMacro(MacroExecutor.MacroParameters mp)
            : base(mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods

        #region ContactDetailsMacro
        /// <summary>
        /// /// Displays contact information (Fax only) based on Role Type
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
        public static MacroExecutor.MacroRetCd Details1(
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.PregnancyMacro.Pregnancy,ProtocolDTs.dll" elementLabel="Pregnancy" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Population" autogenerates="true" toolTip="Pregnancy." shouldRun="true"/>
#endif
			try 
			{
                mp.pba_.setOperation("Contact Details Macro", "Generating information...");

                ContactDetailsMacro macro = null;
                macro = new ContactDetailsMacro(mp);
				macro.preProcess();
                macro.displayFax();
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
                Log.exception(e, "Error in Contact Details Macro");
                mp.inoutRng_.Text = "Contact Details Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		#endregion

        #region Details2
        /// <summary>
        /// Displays contact information based on Role Type
        /// </summary>
        /// <param name="mp"></param>
        /// <returns></returns>
        public static MacroExecutor.MacroRetCd Details2(
            MacroExecutor.MacroParameters mp)
        {
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.PregnancyMacro.Pregnancy,ProtocolDTs.dll" elementLabel="Pregnancy" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Population" autogenerates="true" toolTip="Pregnancy." shouldRun="true"/>
#endif
            try
            {
                mp.pba_.setOperation("Contact Details Macro", "Generating information...");

                ContactDetailsMacro macro = null;
                macro = new ContactDetailsMacro(mp);
                macro.preProcess();
                macro.displaycontactinfo();
                macro.postProcess();
                return macro.macroStatusCode_;
            }
            catch (Exception e)
            {
                Log.exception(e, "Error in Contact Details Macro");
                mp.inoutRng_.Text = "Contact Details Macro: " + e.Message;
            }
            return MacroExecutor.MacroRetCd.Failed;
        }

        #endregion


		#endregion
        public MacrosConfig mc = null;

        public void displayFax()
        {

            string chooserElementPath = this.macroEntry_.getElementPath();
            string fPath = tspdDoc_.getTrialProject().getTemplateDirPath() + "\\dyntmplts\\MacrosConfig.xml";
            mc = new MacrosConfig(fPath, chooserElementPath);


            Word.Range inoutRange = this.startAtBeginningOfParagraph();
            Word.Range wrkRng = inoutRange.Duplicate;

            pba_.updateProgress(1.0);

            string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);
            BusinessObjectMgr bom_ = tspdDoc_.getBom();

            string strRoleType = mc.getMessageByName("contacttype").Text;
            bool hasContact = false;

            ContactEnumerator conEnum = bom_.getContactEnumerator();
            string strFax = "";
            foreach (Contact c in conEnum.getList())
            {
                if (c.getRoleType().ToLower() == strRoleType.ToLower())
                {
                    strFax = c.getFax();
                    hasContact = true;
                    break;  //Exit after first instance (rest are skipped).
                }
            }

            string msg = "";
            if (!hasContact)
            {
                //If no contact with "Med Monitor" Found.
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

            if (strFax == null || strFax.Trim().Length <= 0)
            {
                strFax = "###-###-####";
            }

            msg = mc.getMessageByName("maintext").Text;
            msg = msg.Replace("[[fax]]", strFax);


            mc.setStyle(mc.getMessageByName("maintext").Format.Style, tspdDoc_, wrkRng);
            WordFormatter.FTToWordFormat2(ref wrkRng, msg);
            wrkRng.InsertParagraphAfter();
            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
            // Set outgoing range
            inoutRange.End = wrkRng.End;
            setOutgoingRng(inoutRange);
            wdDoc_.UndoClear();
        }


        public void displaycontactinfo()
        {


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
            }


            Word.Range inoutRange = this.startAtBeginningOfParagraph();
            Word.Range wrkRng = inoutRange.Duplicate;

            pba_.updateProgress(1.0);

            string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);
            BusinessObjectMgr bom_ = tspdDoc_.getBom();

            string strRoleType = mc.getMessageByName("contacttype").Text;


            ContactEnumerator conEnum = bom_.getContactEnumerator();
            string strFax = "";
            string strCROName = "";
            string strPhone = "";
            string strEmail = "";
            bool hasContact = false;

            foreach (Contact c in conEnum.getList())
            {
                if (c.getRoleType().ToLower() == strRoleType.ToLower())
                {
                    strFax = c.getFax();
                    strCROName = c.getActualDisplayValue();
                    strPhone = c.getTel();
                    strEmail = c.getEmail();
                    hasContact = true;
                    break;  //Exit after first instance (rest are skipped).
                }
            }

            string msg = "";
            if (!hasContact)
            {
                //If no contact with "Med Monitor" Found.
                msg = mc.getMessageByName("exception1").Text;
                wrkRng.InsertAfter(msg);
                wrkRng.InsertParagraphAfter();
                wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
                // Set outgoing range
                inoutRange.End = wrkRng.End;
                setOutgoingRng(inoutRange);
                wdDoc_.UndoClear();
            }

            if (strFax == null || strFax.Trim().Length <= 0)
            {
                strFax = "###-###-####";
            }

            if (strPhone == null || strPhone.Trim().Length <= 0)
            {
                strPhone = "###-###-####";
            }

            msg =  mc.getMessageByName("maintext").Text;
            msg = msg.Replace("[[fax]]", strFax);
            msg = msg.Replace("[[croname]]", strCROName);
            msg = msg.Replace("[[phone]]", strPhone);
            msg = msg.Replace("[[email]]", strEmail);



            mc.setStyle(mc.getMessageByName("maintext").Format.Style, tspdDoc_, wrkRng);
            WordFormatter.FTToWordFormat2(ref wrkRng, msg);
            wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
            // Set outgoing range
            inoutRange.End = wrkRng.End;
            setOutgoingRng(inoutRange);
            wdDoc_.UndoClear();
        }

		public override void postProcess()
		{
			// Clean up memory
		}
	}
}
