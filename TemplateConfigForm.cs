using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using System.IO;
//using Aspose.Words;
using System.Xml;
using K12.Data.Configuration;
using System.Diagnostics;
using DevComponents.DotNetBar.Controls;
using FISCA.Presentation.Controls;

namespace OfficialStudentRecordReport2010
{
    public partial class TemplateConfigForm : BaseForm
    {
        private int _useTemplateNumber = 0;

        private bool getOut = false;
        private RadioButton cRB = null;

        private byte[] _buffer1 = null;
        private byte[] _buffer2 = null;
        private byte[] _buffer3 = null;

        private string _base64string1 = null;
        private string _base64string2 = null;
        private string _base64string3 = null;

        private bool _isUpload1 = false;
        private bool _isUpload2 = false;
        private bool _isUpload3 = false;

        private int textScoreOption = 0;

        public TemplateConfigForm(int useTemplateNumber, byte[] buffer1, byte[] buffer2, byte[] buffer3, int[] print, bool over100, string coreSubjectSign, string coreCourseSign, string resitSign, string retakeSign, string schoolYearAdjustSign, string manualAdjustSign, string failedSign, int moralScoreOption)
        {
            InitializeComponent();

            textScoreOption = moralScoreOption;

            chkClassTeacher.Checked = false;
            chkTextScore.Checked = false;
            if (moralScoreOption == 1)
                chkClassTeacher.Checked = true;
            if (moralScoreOption == 2)
                chkTextScore.Checked = true;
            if (moralScoreOption == 3)
            {
                chkClassTeacher.Checked = true;
                chkTextScore.Checked = true;
            }

            _useTemplateNumber = useTemplateNumber;
            switch (_useTemplateNumber)
            {
                case 1:
                    checkBoxX1.Checked = true;
                    break;
                case 2:
                    checkBoxX2.Checked = true;
                    break;
                case 3:
                    checkBoxX3.Checked = true;
                    break;
                case 4:
                    checkBoxX4.Checked = true;
                    break;
                case 5:
                    checkBoxX5.Checked = true;
                    break;
                case 6:
                    checkBoxX6.Checked = true;
                    break;
                default:
                    break;
            }

            _buffer1 = buffer1;
            _buffer2 = buffer2;
            _buffer3 = buffer3;

            cboRecvIdentity.SelectedIndex = print[0];
            cboRecvAddress.SelectedIndex = print[1];
            cboRecvPhone.SelectedIndex = print[2];

            txtCoreSubjectSign.Text = coreSubjectSign;
            txtCoreCourseSign.Text = coreCourseSign;
            txtResitSign.Text = resitSign;
            txtRetakeSign.Text = retakeSign;

            txtFailedSign.Text = failedSign;
            txtSchoolYearAdjustSign.Text = schoolYearAdjustSign;
            txtManualAdjustSign.Text = manualAdjustSign;
        }

        //  檢視「高級中學用學籍表預設範本 (含學年成績)」
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "高中學籍表.xls";
            sfd.Filter = "相容於 Excel 2003 檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    DownloadTemplate(sfd, Properties.Resources.學籍表高中);
                }
                catch (Exception ex)
                {
                    MsgBox.Show(ex.Message);
                }
            }
        }

        //  檢視「高級中學用學籍表自訂範本 (含學年成績)」
        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "自訂高中學籍表.xls";
            sfd.Filter = "相容於 Excel 2003 檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    DownloadTemplate(sfd, _buffer1);
                }
                catch (Exception ex)
                {
                    MsgBox.Show(ex.Message);
                }
            }
        }

        //  上傳「高級中學用學籍表自訂範本 (含學年成績)」
        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "上傳自訂高中學籍表範本";
            ofd.Filter = "相容於 Excel 2003 檔案 (*.xls)|*.xls";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    UploadTemplate(ofd.FileName, ref _isUpload1, ref _base64string1);
                }
                catch (Exception ex)
                {
                    MsgBox.Show(ex.Message);
                }
            }
        }

        //  檢視「職業學校用學籍表預設範本」
        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "高職學籍表.xls";
            sfd.Filter = "相容於 Excel 2003 檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    DownloadTemplate(sfd, Properties.Resources.學籍表高職);
                }
                catch (Exception ex)
                {
                    MsgBox.Show(ex.Message);
                }
            }
        }

        //  檢視「職業學校用學籍表自訂範本」
        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "自訂高職學籍表.xls";
            sfd.Filter = "相容於 Excel 2003 檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    DownloadTemplate(sfd, _buffer2);
                }
                catch (Exception ex)
                {
                    MsgBox.Show(ex.Message);
                }
            }
        }

        //  上傳「職業學校用學籍表自訂範本」
        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "上傳自訂職業學校用學籍表範本";
            ofd.Filter = "相容於 Excel 2003 檔案 (*.xls)|*.xls";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    UploadTemplate(ofd.FileName, ref _isUpload2, ref _base64string2);
                }
                catch (Exception ex)
                {
                    MsgBox.Show(ex.Message);
                }
            }
        }

        //  檢視「進修學校用學籍表預設範本 (含學年成績)」
        private void linkLabel8_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "進校學籍表.xls";
            sfd.Filter = "相容於 Excel 2003 檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    DownloadTemplate(sfd, Properties.Resources.學籍表進校);
                }
                catch (Exception ex)
                {
                    MsgBox.Show(ex.Message);
                }
            }
        }

        //  檢視「進修學校用學籍表自訂範本 (含學年成績)」
        private void linkLabel9_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "自訂進校學籍表.xls";
            sfd.Filter = "相容於 Excel 2003 檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    DownloadTemplate(sfd, _buffer3);
                }
                catch (Exception ex)
                {
                    MsgBox.Show(ex.Message);
                }
            }
        }

        //  上傳「進修學校用學籍表自訂範本 (含學年成績)」
        private void linkLabel7_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "上傳自訂進修學校用學籍表範本";
            ofd.Filter = "相容於 Excel 2003 檔案 (*.xls)|*.xls";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    UploadTemplate(ofd.FileName, ref _isUpload3, ref _base64string3);
                }
                catch (Exception ex)
                {
                    MsgBox.Show(ex.Message);
                }
            }
        }

        //  下載模組
        private void DownloadTemplate(SaveFileDialog sfd, byte[] fileData)
        {
            if ((fileData == null) || (fileData.Length == 0))
            {
                throw new Exception("檔案不存在，無法檢視。");
            }

            try
            {
                FileStream fs = new FileStream(sfd.FileName, FileMode.Create);

                fs.Write(fileData, 0, fileData.Length);
                fs.Close();
                System.Diagnostics.Process.Start(sfd.FileName);
            }
            catch
            {
                throw new Exception("指定路徑無法存取。");
            }
        }

        //  上傳模組
        private void UploadTemplate(string fileName, ref bool uploadIndex, ref string uploadData)
        {
            FileStream fs = new FileStream(fileName, FileMode.Open);

            byte[] tempBuffer = new byte[fs.Length];
            fs.Read(tempBuffer, 0, tempBuffer.Length);

            MemoryStream ms = new MemoryStream(tempBuffer);

            try
            {
                Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();

                wb.Open(ms, Aspose.Cells.FileFormatType.Excel2003);
                wb = null;
            }
            catch
            {
                throw new Exception("此版學籍表範本限用相容於 Excel 2003 檔案。");
            }

            try
            {
                uploadData = Convert.ToBase64String(tempBuffer);
                uploadIndex = true;

                fs.Close();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        //  離開
        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Dispose();
        }

        //  設定
        private void buttonX1_Click(object sender, EventArgs e)
        {
            #region 儲存 Preference

            ConfigData config = K12.Data.School.Configuration["2010學生學籍表"];
            config["TemplateNumber"] = _useTemplateNumber.ToString();

            if (_isUpload1)
                config["CustomizeTemplate1"] = _base64string1;

            if (_isUpload2)
                config["CustomizeTemplate2"] = _base64string2;

            if (_isUpload3)
                config["CustomizeTemplate3"] = _base64string3;

            config["Custodian"] = cboRecvIdentity.SelectedIndex.ToString();
            config["Address"] = cboRecvAddress.SelectedIndex.ToString();
            config["Phone"] = cboRecvPhone.SelectedIndex.ToString();
            config["CoreSubjectSign"] = txtCoreSubjectSign.Text;
            config["CoreCourseSign"] = txtCoreCourseSign.Text;
            config["ResitSign"] = txtResitSign.Text;
            config["RetakeSign"] = txtRetakeSign.Text;
            config["SchoolYearAdjustSign"] = txtSchoolYearAdjustSign.Text;
            config["ManualAdjustSign"] = txtManualAdjustSign.Text;
            config["FailedSign"] = txtFailedSign.Text;
            config["MoralScoreOption"] = textScoreOption.ToString();

            config.Save();

            #endregion

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void checkBoxX_Click(object sender, EventArgs e)
        {
            if (getOut)
                return;

            cRB = (RadioButton)sender;
            getOut = true;

            checkBoxX1.Checked = false;
            checkBoxX2.Checked = false;
            checkBoxX3.Checked = false;
            checkBoxX4.Checked = false;
            checkBoxX5.Checked = false;
            checkBoxX6.Checked = false;

            cRB.Checked = true;
            
            getOut = false;

            _useTemplateNumber = Convert.ToInt32(((RadioButton)sender).Name.Substring((((RadioButton)sender).Name.Length - 1), 1));
        }

        private void chkClassTeacher_CheckedChanged(object sender, EventArgs e)
        {
            SumTextScoreOption();
        }

        private void chkTextScore_CheckedChanged(object sender, EventArgs e)
        {
            SumTextScoreOption();
        }

        private void SumTextScoreOption()
        {
            int no = 0;
            if (chkClassTeacher.Checked == true)
                no += 1;

            if (chkTextScore.Checked == true)
                no += 2;

            textScoreOption = no;
        }
    }
}