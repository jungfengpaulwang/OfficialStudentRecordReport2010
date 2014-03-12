using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using System.Xml;
//using SmartSchool.Customization.Data;
using K12.Data.Configuration;

namespace OfficialStudentRecordReport2010
{
    public partial class TagConfig : FISCA.Presentation.Controls.BaseForm
    {
        XmlElement _categories;
        Dictionary<string, string> _tagList;

        ConfigData config;

        public TagConfig(XmlElement categories, Dictionary<string, string> tagList)
        {
            InitializeComponent();
            
            config = K12.Data.School.Configuration["2010學生學籍表"];  

            _categories = categories;
            _tagList = tagList;
        }

        private void TagConfig_Load(object sender, EventArgs e)
        {
            foreach (XmlElement tag in _categories.SelectNodes("Tag"))
            {
                ListViewItem item = new ListViewItem();
                string prefix = tag.SelectSingleNode("Prefix").InnerText;
                string name = tag.SelectSingleNode("Name").InnerText;
                string id = tag.GetAttribute("ID");
                if (!string.IsNullOrEmpty(prefix))
                    name = prefix + ":" + name;
                item.Text = name;
                item.Tag = id;
                if (_tagList.ContainsKey(id))
                    item.Checked = true;
                listViewEx1.Items.Add(item);
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            #region 儲存 Preference            
           
            XmlDocument xDoc = new XmlDocument();
            XmlElement Tags = xDoc.CreateElement("Tags");
            foreach (ListViewItem item in listViewEx1.Items)
            {
                if (item.Checked == true)
                {
                    XmlElement tag = Tags.OwnerDocument.CreateElement("Tag");
                    tag.SetAttribute("ID", item.Tag as string);
                    tag.SetAttribute("FullName", item.Text);
                    Tags.AppendChild(tag);
                }
            }

            if (Tags.HasChildNodes)
            {
                config["Tags"] = Tags.InnerXml;
            }

            #endregion

            this.DialogResult = DialogResult.OK;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
    }
}