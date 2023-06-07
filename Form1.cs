using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NightShiftExcelHelper {
    public partial class NightShiftExcel : Form {
        public NightShiftExcel() {
            InitializeComponent();
            msg.Text = "";
            string base64 = NetHelper.GetCode();
            base64 = base64.Replace("data:image/png;base64,", "")
                .Replace("data:image/jgp;base64,", "").Replace("data:image/jpg;base64,", "")
                .Replace("data:image/jpeg;base64,", "");//将base64头部信息替换
            byte[] bytes = Convert.FromBase64String(base64);
            MemoryStream memStream = new MemoryStream(bytes);
            Image mImage = Image.FromStream(memStream);
            code.Image = mImage;
            
            msg.Text = "自由探索".ToCharArray()[0].ToString();
        }

        private void rzoaBtn_Click(object sender, EventArgs e) {
            FileChoose("rzoa");
            if (ExcelHelper.excelDic.ContainsKey("rzoa")) {
                rzoaTxt.Text = ExcelHelper.excelDic["rzoa"];
            }
        }

        private void erpBtn_Click(object sender, EventArgs e) {
            FileChoose("erp");
            if (ExcelHelper.excelDic.ContainsKey("erp")) {
                erpTxt.Text = ExcelHelper.excelDic["erp"];
            }
        }

        private void portalBtn_Click(object sender, EventArgs e) {
            FileChoose("portal");
            if (ExcelHelper.excelDic.ContainsKey("portal")) {
                portalTxt.Text = ExcelHelper.excelDic["portal"];
            }
        }

        private void FileChoose(String key) {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;      //该值确定是否可以选择多个文件
            openFileDialog.Title = "请选择文件";     //弹窗的标题
            //openFileDialog.InitialDirectory = "D:\\";       //默认打开的文件夹的位置
            openFileDialog.Filter = "MicroSoft Excel文件(*.xlsx)|*.xlsx|所有文件(*.*)|*.*";       //筛选文件
            openFileDialog.ShowHelp = true;     //是否显示“帮助”按钮
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            if (openFileDialog.ShowDialog() == DialogResult.OK) {
                if (ExcelHelper.excelDic.ContainsKey(key)) {
                    ExcelHelper.excelDic[key] = openFileDialog.FileName;
                } else {
                    ExcelHelper.excelDic.Add(key, openFileDialog.FileName);
                }
            }
        }

        private void checkBtn_Click(object sender, EventArgs e) {
            //msg.Text = "";
            //try {
            //    ExcelHelper.Check();
            //    msg.Text = "检查完成！";
            //}
            //catch (Exception exp) {
            //    msg.Text = exp.Message;
            //}

            if (codeTxt.Text.Trim() == "") {
                return;
            }
            msg.Text = "登录中";
            if (NetHelper.Login(codeTxt.Text.Trim())) {
                code.Visible = false;
                codeTxt.Enabled = false;
                codeTxt.Visible = false;
            } else {
                msg.Text = "登录失败";
            }

            msg.Text = "认证中";
            msg.Refresh();
            NetHelper.Authentication();

            msg.Text = "弱口令信息统计中：RZOA";
            msg.Refresh();
            NetHelper.GetWPList("10.200.24.11");
            msg.Text = "弱口令信息统计中：ERP";
            msg.Refresh();
            NetHelper.GetWPList("10.200.35.56");

            msg.Text = "安全事件统计中：高危";
            msg.Refresh();
            NetHelper.GetSIList("High");
            msg.Text = "安全事件统计中：中危";
            msg.Refresh();
            NetHelper.GetSIList("Medium");
            msg.Text = "安全事件统计中：低危";
            msg.Refresh();
            NetHelper.GetSIList("Low");

            msg.Text = "风险资产统计中：内网";
            msg.Refresh();
            NetHelper.GetRAList("内网");
            msg.Text = "风险资产统计中：外网";
            msg.Refresh();
            NetHelper.GetRAList("外网");
            msg.Text = "风险资产统计中：工控网";
            msg.Refresh();
            NetHelper.GetRAList("工控网");
            msg.Text = "风险资产统计中：其他";
            msg.Refresh();
            NetHelper.GetRAList("其他");

            msg.Text = "弱口令信息测试中：RZOA";
            msg.Refresh();
            NetHelper.RzoaRequset();
            msg.Text = "弱口令信息测试中：ERP";
            msg.Refresh();
            NetHelper.ErpRequset();

            //NetHelper.Test();
            msg.Text = "信息安全相关网站测试中";
            msg.Refresh();
            msg.Text = "生成ERP弱口令表格";
            msg.Refresh();
            ExcelHelper.FillWPSheet(false);
            msg.Text = "生成RZOA弱口令表格";
            msg.Refresh();
            ExcelHelper.FillWPSheet(true);
            msg.Text = "生成信息安全表格";
            msg.Refresh();
            ExcelHelper.FillNSSheet();

            msg.Text = "完成！\n又是偷懒成功的一天呢";

            Timer timer = new Timer();
            Application.Exit();
        }

        private void rzoaCanBtn_Click(object sender, EventArgs e) {
            if (ExcelHelper.excelDic.ContainsKey("rzoa")) {
                ExcelHelper.excelDic.Remove("rzoa");
                rzoaTxt.Text = "";
            }
        }

        private void erpCanBtn_Click(object sender, EventArgs e) {
            if (ExcelHelper.excelDic.ContainsKey("erp")) {
                ExcelHelper.excelDic.Remove("erp");
                erpTxt.Text = "";
            }
        }

        private void portalCanBtn_Click(object sender, EventArgs e) {
            if (ExcelHelper.excelDic.ContainsKey("portal")) {
                ExcelHelper.excelDic.Remove("portal");
                portalTxt.Text = "";
            }
        }

        private void button1_Click(object sender, EventArgs e) {
            //msg.Text = NetHelper.RzoaRequset("liuyq", "psss").ToString();
            //msg.Text = NetHelper.ErpRequset("003308", "199961").ToString();
            ExcelHelper.Test();
        }
    }
}
