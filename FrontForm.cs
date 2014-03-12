using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Aspose.Cells;
using FISCA.Presentation.Controls;
using K12.Data.Configuration;

namespace OfficialStudentRecordReport2010
{
    public partial class FrontForm : FISCA.Presentation.Controls.BaseForm
    {
        private MemoryStream _defaultTemplate1 = new MemoryStream(Properties.Resources.���y����);  //  �w�]�������y��d��
        private MemoryStream _template1 = null;                                                     //  �ۭq�������y��d��
        private MemoryStream _defaultTemplate2 = new MemoryStream(Properties.Resources.���y��¾);  //  �w�]��¾���y��d��
        private MemoryStream _template2 = null;                                                     //  �ۭq��¾���y��d��
        private MemoryStream _defaultTemplate3 = new MemoryStream(Properties.Resources.���y��i��);  //  �w�]�i�վ��y��d��
        private MemoryStream _template3 = null;                                                     //  �ۭq�i�վ��y��d��

        private int _useTemplateNumber = 0;
        private int moralScoreOption = 0;
        private byte[] _buffer1 = null;
        private byte[] _buffer2 = null;
        private byte[] _buffer3 = null;

        ConfigData config;

        private int _optSaveFileType = 0;
        private bool getOut = false;

        public int SaveFileType
        {
            get { return _optSaveFileType; }
        }

        public int TemplateNumber
        {
            get { return _useTemplateNumber; }
        }

        public int MoralScoreOption
        {
            get { return moralScoreOption; }
        }
        
        public MemoryStream Template
        {
            get
            {
                switch (_useTemplateNumber)
                {
                    case 1:
                        return _defaultTemplate1;
                    case 2:
                        return _template1;
                    case 3:
                        return _defaultTemplate2;
                    case 4:
                        return _template2;
                    case 5:
                        return _defaultTemplate3;
                    case 6:
                        return _template3;
                    default:
                        return new MemoryStream();
                }
            }
        }

        private bool _error2 = false;

        private string _text1;
        private string _text2;
        private string _text3;

        public string Text1
        {
            get { return _text1; }
        }

        public string Text2
        {
            get { return _text2; }
        }

        public string Text3
        {
            get { return _text3; }
        }

        private int _custodian = 0;
        private int _address = 0;
        private int _phone = 0;

        public int Custodian
        {
            get { return _custodian; }
        }

        public int Address
        {
            get { return _address; }
        }

        public int Phone
        {
            get { return _phone; }
        }

        private bool _over100 = false;
        public bool AllowMoralScoreOver100
        {
            get { return _over100; }
        }

        #region �ۭq�Х�

        /// <summary>
        /// ��������Ш|�֤߬�ؼХ�
        /// </summary>
        private string _sign_core_subject = "";
        public string SignCoreSubject
        {
            get { return _sign_core_subject; }
        }

        /// <summary>
        /// ��X�����ǵ{�֤߽ҵ{�Х�
        /// </summary>
        private string _sign_core_course = "";
        public string SignCoreCourse
        {
            get { return _sign_core_course; }
        }

        /// <summary>
        /// �����o�Ǥ��Х�
        /// </summary>
        private string _sign_failed = "*";
        public string SignFailed
        {
            get { return _sign_failed; }
        }

        /// <summary>
        /// �Ǧ~�վ㦨�Z�Х�
        /// </summary>
        private string _sign_school_year_adjust = "";
        public string SignSchoolYearAdjust
        {
            get { return _sign_school_year_adjust; }
        }

        /// <summary>
        /// ��ʽվ㦨�Z�Х�
        /// </summary>
        private string _sign_manual_adjust = "";
        public string SignManualAdjust
        {
            get { return _sign_manual_adjust; }
        }

        /// <summary>
        /// �ɦҦ��Z�Х�
        /// </summary>
        private string _sign_resit = "";
        public string SignResit
        {
            get { return _sign_resit; }
        }

        /// <summary>
        /// ���צ��Z�Х�
        /// </summary>
        private string _sign_retake = "";
        public string SignRetake
        {
            get { return _sign_retake; }
        }

        #endregion

        private Dictionary<string, string> _tagList = new Dictionary<string, string>();
        public Dictionary<string, string> TagList
        {
            get { return _tagList; }
        }

        public FrontForm()
        {
            InitializeComponent();

            //string DALMessage = "�y";

            //foreach (Assembly Assembly in AppDomain.CurrentDomain.GetAssemblies().Where(x => x.GetName().Name.Equals("OfficialStudentRecordReport2010")))
            //    DALMessage += "�������G" + Assembly.GetName().Version + " ";

            //DALMessage += "�z";

            //this.Text += DALMessage;
           
            LoadPreference();

            textBoxX1.Text = this.Text1;
            textBoxX2.Text = this.Text2;

            //string path = Path.Combine(Application.StartupPath, @"Reports\���y��");
            //if (!Directory.Exists(path))
            //    Directory.CreateDirectory(path);

            //ReportDirectory.Text = path;
        }

        private void LoadPreference()
        {
            #region Ū�� Preference

            config = K12.Data.School.Configuration["2010�ǥ;��y��"];   
            if (config != null)
            {
                int no = 0;
                int.TryParse(config["TemplateNumber"], out no);
                _useTemplateNumber = no;

                string customize1 = config["CustomizeTemplate1"];
                string customize2 = config["CustomizeTemplate2"];
                string customize3 = config["CustomizeTemplate3"];
                _text1 = config["Word"];
                _text2 = config["Number"];
                //_text3 = config["ReportDirectory"];                

                if (!string.IsNullOrEmpty(customize1))
                {
                    _buffer1 = Convert.FromBase64String(customize1);
                    _template1 = new MemoryStream(_buffer1);
                }

                if (!string.IsNullOrEmpty(customize2))
                {
                    _buffer2 = Convert.FromBase64String(customize2);
                    _template2 = new MemoryStream(_buffer2);
                }

                if (!string.IsNullOrEmpty(customize3))
                {
                    _buffer3 = Convert.FromBase64String(customize3);
                    _template3 = new MemoryStream(_buffer3);
                }

                _custodian = 0;
                int.TryParse(config["Custodian"], out _custodian);

                _address = 0;
                int.TryParse(config["Address"], out _address);

                _phone = 0; 
                int.TryParse(config["Phone"], out _phone);

                moralScoreOption = 0;
                int.TryParse(config["MoralScoreOption"], out moralScoreOption);

                _optSaveFileType = 0;
                int.TryParse(config["SaveFileType"], out _optSaveFileType);

                switch (_optSaveFileType)
                {
                    case 1:
                        this.checkBoxX1.Checked = true;
                        break;
                    case 2:
                        this.checkBoxX2.Checked = true;
                        break;
                    case 3:
                        this.checkBoxX3.Checked = true;
                        break;
                }

                _sign_core_subject = config["CoreSubjectSign"];                
                _sign_core_course = config["CoreCourseSign"];
                _sign_resit = config["ResitSign"];
                _sign_retake = config["RetakeSign"];
                _sign_failed = config["FailedSign"];
                _sign_school_year_adjust = config["SchoolYearAdjustSign"];
                _sign_manual_adjust = config["ManualAdjustSign"];

                //string xml = config["Tags"];

                //if (String.IsNullOrEmpty(xml))
                //    return;

                //XmlDocument xDoc = new XmlDocument();
                //xDoc.LoadXml("<root>" + xml + "</root>");

                //foreach (XmlElement tag in xDoc.DocumentElement.SelectNodes("Tag"))
                //{
                //    if (!_tagList.ContainsKey(tag.GetAttribute("ID")))
                //        _tagList.Add(tag.GetAttribute("ID"), tag.GetAttribute("FullName"));
                //}
            }
            #endregion
        }

        private void SavePreference()
        {
            #region �x�sPreference

            if (config != null)
            {
                config["Word"] = _text1;
                config["Number"] = _text2;
                //config["ReportDirectory"] = _text3;
                config["SaveFileType"] = _optSaveFileType.ToString();
            }
            config.Save();

            #endregion
        }
        
        private void Print_Click(object sender, EventArgs e)
        {
            if (_error2)
                return;

            //if (!System.IO.Directory.Exists(this.ReportDirectory.Text))
            //{
            //    MsgBox.Show("�п�ܾ��y���X���|�I");
            //    return;
            //}

            if (this.TemplateNumber == 0)
            {
                MsgBox.Show("�Ы��u�d���]�w�v�A��ܼ˪��I");
                return;
            }

            Workbook wb = new Workbook();
            try
            {
                wb.Open(this.Template, FileFormatType.Excel2003);
            }
            catch
            {
                MsgBox.Show("�ФW�Ǭۮe�� Excel 2003 ���d���ɮסI");
                return;
            }
            finally
            {
                wb = null;
            }

            if (this.SaveFileType == 0)
            {                
                MsgBox.Show("�г]�w�u���y��s�ɿﶵ�v");
                return;
            }

            _text3 = "";
            System.Windows.Forms.FolderBrowserDialog folder = new FolderBrowserDialog();

            if (folder.ShowDialog() == DialogResult.OK)
            {
                _text3 = folder.SelectedPath;
            }

            _text1 = textBoxX1.Text;
            _text2 = textBoxX2.Text;

            SavePreference();

            if (Directory.Exists(_text3))
            {
                //MsgBox.Show("�п�ܡu���y��s�ɥؿ��v");
                //return;
                this.DialogResult = DialogResult.OK;
            }
        }

        private void textBoxX2_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBoxX2.Text))
            {
                _error2 = false;
                int a = 0;

                foreach (char var in textBoxX2.Text.ToCharArray())
                {
                    if (!int.TryParse(var.ToString(), out a))
                    {
                        _error2 = true;
                        break;
                    }
                }

                if (_error2)
                    errorProvider2.SetError(textBoxX2, "�榡���Ʀr");
                else
                    errorProvider2.Clear();
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SavePreference();

            TemplateConfigForm form = new TemplateConfigForm(
                _useTemplateNumber,
                _buffer1,
                _buffer2, 
                _buffer3,
                new int[] { _custodian, _address, _phone },
                _over100,
                _sign_core_subject, 
                _sign_core_course,
                _sign_resit,
                _sign_retake,
                _sign_school_year_adjust,
                _sign_manual_adjust,
                _sign_failed,
                moralScoreOption);

            if (form.ShowDialog() == DialogResult.OK)
            {
                LoadPreference();
            }
        }

        //  �]�w���Z���� 
        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //SmartSchool.Customization.Data.SystemInformation.getField("StudentCategories");
            //XmlElement categories = SmartSchool.Customization.Data.SystemInformation.Fields["StudentCategories"] as XmlElement;
            //TagConfig form = new TagConfig(categories, _tagList);
            //if (form.ShowDialog() == DialogResult.OK)
            //{
            //    LoadPreference();
            //}
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Dispose();
        }

        private void checkBoxX_CheckedChanged(object sender, EventArgs e)
        {
            //if (getOut)
            //    return;

            //cRB = (RadioButton)sender;
            //getOut = true;

            //checkBoxX1.Checked = false;
            //checkBoxX2.Checked = false;
            //checkBoxX3.Checked = false;

            //cRB.Checked = true;

            //getOut = false;

            _optSaveFileType = Convert.ToInt32(((System.Windows.Forms.RadioButton)sender).Name.Substring((((System.Windows.Forms.RadioButton)sender).Name.Length - 1), 1));
        }

        //private void ChangePath_Click(object sender, EventArgs e)
        //{
        //    //System.Windows.Forms.FolderBrowserDialog folder = new FolderBrowserDialog();

        //    //if (folder.ShowDialog() == DialogResult.OK)
        //    //    this.ReportDirectory.Text = folder.SelectedPath;
        //}
    }
}