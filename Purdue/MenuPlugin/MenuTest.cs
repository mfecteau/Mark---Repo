using System;
using System.IO;
using System.Collections;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

using Tspd.Context;
using Tspd.Tspddoc;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;
using Tspd.FormBase;

using System.Xml;
using System.Windows.Forms;


namespace TspdCfg.FastTrack.PlugIn
{

    #region TVMapper

    public class MenuTVMapper : TSDAddinIF
    {
        PublicMenuEventArgs _menutvmapper =
            new PublicMenuEventArgs(
                PublicMenuEventArgs.EventType.DesignGuideMenuEvent,
                "Task-Event Details Mapping",
                typeof(MenuTVMapper).Name);

        #region TSDAddinIF Members

        public void InitializeAddin(DesignerContext cm)
        {
            cm.getPublicEventMgr().MenuEvents += new EventHandler(handleMenuEvents);

            cm.getPublicEventMgr().addMenuItem(_menutvmapper);
        }

        #endregion

        public void handleMenuEvents(object source, EventArgs args)
        {
            PublicMenuEventArgs margs = args as PublicMenuEventArgs;

            DesignerContext cm = DesignerContext.getInstance();
            DesignerDocBase doc = cm.getActiveBaseDocument();
            if (margs == _menutvmapper)
            {                
                
                LoadForm(_menutvmapper.MenuClass + "."+ _menutvmapper.MenuItemName);
                return;
            }
        }


        private void LoadForm(string elementPath)
        {
            frmTVMapper formobject = new frmTVMapper(elementPath);
            formobject.ShowDialog();
        }

        private void changeScheduleName()
        {
            DesignerContext cm = DesignerContext.getInstance();
            DesignerDocBase doc = cm.getActiveBaseDocument();
            SOAEnumerator soaEnum = doc.getBom().getAllSchedules();

            while (soaEnum.MoveNext())
            {
                SOA soa = soaEnum.getCurrent();

                string soaName = soa.getName();
                int paren = soaName.IndexOf('(');

                if (paren != -1)
                {
                    soaName = soaName.Substring(0, paren);
                }

                soaName += "(" + DateTime.Now.ToUniversalTime().ToString() + ")";
                soa.setName(soaName);
            }
        }
    }

    #endregion

    # region TaskSequencer
    public class TaskSequence : TSDAddinIF
    {
        PublicMenuEventArgs _menuTskSeq =
            new PublicMenuEventArgs(
                PublicMenuEventArgs.EventType.DesignGuideMenuEvent,
                "Task Sequencing",
                typeof(TaskSequence).Name);

        #region TSDAddinIF Members

        public void InitializeAddin(DesignerContext cm)
        {
            cm.getPublicEventMgr().MenuEvents += new EventHandler(handleMenuEvents);

            cm.getPublicEventMgr().addMenuItem(_menuTskSeq);
        }

        #endregion

        public void handleMenuEvents(object source, EventArgs args)
        {
            PublicMenuEventArgs margs = args as PublicMenuEventArgs;

            DesignerContext cm = DesignerContext.getInstance();
            DesignerDocBase doc = cm.getActiveBaseDocument();
            if (margs == _menuTskSeq)
            {

                LoadForm(_menuTskSeq.MenuClass + "." + _menuTskSeq.MenuItemName);
                return;
            }
        }


        private void LoadForm(string elementPath)
        {
            frmTaskSeq formobject = new frmTaskSeq(elementPath);
            formobject.ShowDialog();
        }

      
    }
    #endregion




}