using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections;
using MSXML2;

using Tspd.Businessobject;
using Tspd.Tspddoc;
using Tspd.Icp;
using Tspd.Utilities;
using Tspd.Rules;
using Tspd.Context;
using WorX;
using Word = Microsoft.Office.Interop.Word;


namespace TspdCfg.FastTrack.Rules
{
    public class StyleCheckRule : ITSPDRule
    {
#if false
<rule id="Advisory05" type="CSHARP" displayName="StyleCheckRule" source="TspdCfg.FastTrack.Rules.ReferenceCheckRule.,FTRules.dll" categories="StyleAdv" debug="false"/>
#endif
        public static Hashtable Sel_Section = new Hashtable();
        public static readonly string MY_ID = "StyleCheckRule";
        private ArrayList finalResult = new ArrayList();
        private static string _ruleId = "";
        public TspdDocument doc = null;
        private static bool _debug = false;

        public void Init(string ruleId, bool debug)
        {
            _ruleId = ruleId;
            _debug = debug;
        }

        public string AdvisoryPrefix
        {
            get { return MY_ID; }
        }

        public ICollection Run()
        {
            ArrayList advisories = new ArrayList();
            try
            {
               
                ArrayList docSec = new ArrayList();

                ContextManager ctx = ContextManager.getInstance();
                doc = ctx.getActiveDocument();
                BusinessObjectMgr bom = doc.getBom();


                finalResult.Clear();
                docSec = CheckStyles(doc);
                string smsg = "";

                foreach (string str in docSec)
                {
                    smsg = str;
                    RuleAdvisory adv = new RuleAdvisory(
                        _ruleId, MY_ID + str, smsg);
                    advisories.Add(adv);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return advisories;
        }


        public ArrayList CheckStyles(TspdDocument currDoc)
        {
            /*This method will check if all Styles used in Trial are from standard style in template and have not changed any properties.
             Also, it checks for all style in use are from standard set of stlyes which exists in Protocol template.*/

            string fPath = currDoc.getTrialProject().getTemplateDirPath() + "\\worx\\Style.doc";
            object falsch = false;
            object missing = System.Reflection.Missing.Value;
            object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
            object orgFormat = Word.WdOriginalFormat.wdWordDocument;
         //   object password = AppUtilities.parolaChiave();
            object truth = true;
            object filePath = null;

            try
            {
                if (System.IO.File.Exists(fPath))
                {
                    //If file exists get all the styles in to an Enum.
                    filePath = fPath;
                    //Word.Application wdApp_ = currDoc.getActiveWordDocument().Application;
                    Word.Application App_ = new Word.Application();

                    Word.Document templateDoc_ = App_.Documents.Open(ref filePath, ref falsch, ref falsch, ref falsch, ref missing,
                   ref missing, ref missing, ref missing, ref missing,
                   ref missing, ref missing, ref falsch, ref missing, ref missing, ref missing, ref missing);

                    IEnumerator tmpltstylesEnum = templateDoc_.Styles.GetEnumerator();
                    Hashtable tmpltList = new Hashtable();
                    while (tmpltstylesEnum.MoveNext())
                    {
                        {
                            Word.Style currStyle = (Word.Style)tmpltstylesEnum.Current;
                            tmpltList.Add(currStyle.NameLocal.ToString(), currStyle);
                        }
                    }

                    IEnumerator currstylesEnum = currDoc.getActiveWordDocument().Styles.GetEnumerator();
                    Hashtable currList = new Hashtable();
                    while (currstylesEnum.MoveNext())
                    {
                        {
                            Word.Style currStyle = (Word.Style)currstylesEnum.Current;
                            if (currStyle.InUse)
                            {
                                currList.Add(currStyle.NameLocal.ToString(), currStyle);
                            }
                        }
                    }
                    CompareStyles(tmpltList, currList);
                   // templateDoc_.Close
                    templateDoc_.Saved = true;
                    templateDoc_.Close(ref falsch, ref missing, ref missing);
                    //App1_.Quit(ref falsch, ref missing, ref missing);
                }
                //else
                //{
                //    MessageBox.Show("Style document is missing. Please contact your configuration administrator!", "Style Check Advisory", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //}

            }
            catch (Exception ex)
            {
                Log.exception(ex, ex.Message);
                MessageBox.Show(ex.ToString());
            }


            return finalResult;
        }
           public void CompareStyles(Hashtable tmplStyles, Hashtable currStyles)
        {
            /*This method will compare the styles(In Use) from current document against the document tempalte. 
             *  All TSPD Styles will be ignored
             *  
             */

            try
            {
                object key_StyleName = null;
                object val_StyleObj = null;
                string sourceDesc ="", DestDesc="";

                ICollection keyColl = currStyles.Keys;

                foreach (string s in keyColl)
                {
                    Word.Style currStyle = (Word.Style) currStyles[s];

                    if (currStyle.InUse && currStyle.Type != Word.WdStyleType.wdStyleTypeCharacter)
                    {
                        if (!currStyle.NameLocal.ToString().StartsWith("tspd"))
                        {
                            if (tmplStyles.ContainsKey(currStyle.NameLocal.ToString()))
                            {
                                key_StyleName = currStyle.NameLocal.ToString();
                                val_StyleObj = tmplStyles[key_StyleName];
                                Word.Style hash_Style = (Word.Style)val_StyleObj;

                                //Somehow Word linked is in the description. Removing it 
                                sourceDesc = hash_Style.Description;
                                sourceDesc = sourceDesc.Replace("Linked, ", "");
                                sourceDesc = sourceDesc.Replace("Automatically update, ", "");
                                sourceDesc = sourceDesc.Replace("Hide until used, ", "");
                                sourceDesc = sourceDesc.Replace("Hidden, ", "");
                             

                                DestDesc = currStyle.Description;
                                DestDesc = DestDesc.Replace("Linked, ", "");
                                DestDesc = DestDesc.Replace("Automatically update, ", "");
                                DestDesc = DestDesc.Replace("Hide until used, ", "");
                                DestDesc = DestDesc.Replace("Hidden, ", "");
                             

                                //if (!hash_Style.Description.Equals(currStyle.Description))
                                if(!sourceDesc.Equals(DestDesc))
                                {
                                    finalResult.Add("Style " + currStyle.NameLocal.ToString() + " is a standard style, but has been altered from the standard definition.");
                                    //finalResult.Add("STYLE 1 " + hash_Style.Description);
                                    //finalResult.Add("Current STYLE 2 " + currStyle.Description);
                                }
                            }
                            else
                            {
                                finalResult.Add("Style " + currStyle.NameLocal.ToString() + " is not a standard style.");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

              public void CompareStyles(IEnumerator tmplStyles, IEnumerator currStyles)
        {
            /*This method will compare the styles(In Use) from current document against the document tempalte. 
             *  All TSPD Styles will be ignored
             *  
             */

            try
            {
                while (currStyles.MoveNext())
                {
                   Word.Style currStyle = (Word.Style) currStyles.Current;
                   if (currStyle.InUse)
                   {
                       if (!currStyle.NameLocal.ToString().StartsWith("tspd"))
                       {
                           tmplStyles.Reset();
                           int foundStyle = StyleInTemplate(tmplStyles, currStyle);
                           if (foundStyle == 2)
                           {
                               finalResult.Add("Style " + currStyle.NameLocal.ToString() + " is a standard style, but has been altered from the standard definition.");
                           }
                           else if (foundStyle == 0)
                           {
                               finalResult.Add("Style " + currStyle.NameLocal.ToString() + " is not a standard style.");
                           }
                       }
                   } 
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
              private int StyleInTemplate(IEnumerator tmplStyle, Word.Style sel_Style)
              {
                  /* Function will return integer flag. 
                   * flag = 0 -- Does not exist in template.
                   * flag =1  -- Exist's in document, and is same in both original and template
                   * flag =2 --> Exists in document, but modified in Trial document.
                   * */


                  int flag = 0;
                  while (tmplStyle.MoveNext())
                  {
                      Word.Style t_Style = (Word.Style)tmplStyle.Current;
                      if (t_Style.NameLocal.ToString().Equals(sel_Style.NameLocal.ToString()))
                      {
                          flag = 1;

                          if (!t_Style.Description.Equals(sel_Style.Description))
                          // if (t_Style.Equals(sel_Style))
                          {
                              flag = 2;
                              return flag;
                          }
                          return flag;
                      }
                  }
                  return flag;
              }
    }
}

