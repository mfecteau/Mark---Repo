using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace TspdCfg.Purdue.DynTmplts
{
    public partial class InsertDocSection : Form
    {
        public string _sectionName;
        public InsertDocSection()
        {
            InitializeComponent();
        }

        public void ParseNodeListforSectionSelection(XmlNodeList xiNodelist)
        {
            cmbSectionselection.Items.Clear();
            foreach (XmlNode macroNode in xiNodelist)
            {
                if (macroNode.Attributes.GetNamedItem("hidden").Value.ToLower() == "false")
                {
                    cmbSectionselection.Items.Add(macroNode.Attributes.GetNamedItem("name").Value);
                }
            }
            this.ShowDialog();
        }

        private void cmdOK_Click(object sender, EventArgs e)
        {
            if (cmbSectionselection.SelectedIndex >= 0)
            {
                _sectionName = cmbSectionselection.SelectedItem.ToString();
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
            }
            else
            {
                _sectionName = "";
                this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            }
        }

     

        

    }
}
