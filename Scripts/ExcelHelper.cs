using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace NightShiftExcelHelper {
    class ExcelHelper {

        //模板路径
        private static readonly string rzoaSheet = "./RZOA告警列表.xlsx";
        private static readonly string erpSheet = "./ERP告警列表.xlsx";
        private static readonly string nsSheet = "./信息安全网络安全监控报表.xlsx";

        public static Dictionary<String, String> excelDic = new Dictionary<string, string>();

        public static string time = String.Format("{0:'_'yyyyMMdd}", DateTime.Now);
        private static readonly string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        public static void FillWPSheet(bool isOA) {

            string filePath;
            List<ResponseData.WPData.WPInfo> wpInfos;
            FileInfo newFile;
            if (isOA) {
                filePath = rzoaSheet;
                wpInfos = NetHelper.rzoaList;
                newFile = new FileInfo(desktop + "/RZOA告警列表_" + time + ".xlsx");
            }
            else {
                filePath = erpSheet;
                wpInfos = NetHelper.erpList;
                newFile = new FileInfo(desktop + "/ERP告警列表_" + time + ".xlsx");
            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;    // 添加这行代码后不会报ExcelPackage错误
            using (var p = new ExcelPackage()) { }
            FileInfo existingFile = new FileInfo(filePath);
            ExcelPackage package = new ExcelPackage(existingFile);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

            for (int i = 0; i < wpInfos.Count; i++) {
                
                worksheet.Cells[i + 2, 1].Value = wpInfos[i].srcAddress;
                worksheet.Cells[i + 2, 2].Value = wpInfos[i].destAddress;
                worksheet.Cells[i + 2, 3].Value = wpInfos[i].alarmName;
                worksheet.Cells[i + 2, 4].Value = wpInfos[i].subCategory;
                worksheet.Cells[i + 2, 5].Value = wpInfos[i].destPort;
                worksheet.Cells[i + 2, 6].Value = wpInfos[i].threatSeverity;
                worksheet.Cells[i + 2, 7].Value = wpInfos[i].alarmResults;
                worksheet.Cells[i + 2, 8].Value = wpInfos[i].collectorReceiptTime;
                worksheet.Cells[i + 2, 9].Value = wpInfos[i].srcUserName;
                worksheet.Cells[i + 2, 10].Value = wpInfos[i].passwd;
                worksheet.Cells[i + 2, 11].Value = wpInfos[i].testResult;
            }

            package.SaveAs(newFile);
            package.Dispose();
        }

        public static void FillNSSheet() {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            FileInfo newFile = new FileInfo(desktop + "/信息安全网络安全监控报表_" + time + ".xlsx");

            FileInfo existingFile = new FileInfo(nsSheet);
            ExcelPackage package = new ExcelPackage(existingFile);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

            //数量
            worksheet.Cells[4, 4].Value = NetHelper.highSIList.Count;
            worksheet.Cells[5, 4].Value = NetHelper.mediumSIList.Count;
            worksheet.Cells[6, 4].Value = NetHelper.lowSIList.Count;
            worksheet.Cells[7, 4].Value = NetHelper.exNetList.Count;
            worksheet.Cells[8, 4].Value = NetHelper.inNetList.Count;
            worksheet.Cells[9, 4].Value = NetHelper.icNetList.Count;
            worksheet.Cells[10, 4].Value = NetHelper.otherNetList.Count;

            worksheet.Cells[1, 1].Value +=
                DateTime.Now.Month + "月" + (DateTime.Now.Day - 1) + "日-" 
                + DateTime.Now.Month + "月" + DateTime.Now.Day + "日)";
            //填写高危事件
            foreach (ResponseData.SIData.SIInfo siInfo in NetHelper.highSIList) {

                worksheet.Cells[4, 5].Value += siInfo.incidentName + ":";
                for (int i = 0; i < siInfo.evidence.body.Count; i++) {
                    worksheet.Cells[4, 5].Value += siInfo.evidence.body[i].srcAddress + "、";
                }

                worksheet.Cells[4, 5].Value += "\n";
            }

            //填写中危事件
            foreach (ResponseData.SIData.SIInfo siInfo in NetHelper.mediumSIList) {

                worksheet.Cells[5, 5].Value += siInfo.incidentName + ":";
                for (int i = 0; i < siInfo.evidence.body.Count; i++) {
                    worksheet.Cells[5, 5].Value += siInfo.evidence.body[i].srcAddress + "、";
                }

                worksheet.Cells[5, 5].Value += "\n";
            }

            //填写低危事件
            foreach (ResponseData.SIData.SIInfo siInfo in NetHelper.lowSIList) {

                worksheet.Cells[6, 5].Value += siInfo.incidentName + ":";
                for (int i = 0; i < siInfo.evidence.body.Count; i++) {
                    worksheet.Cells[6, 5].Value += siInfo.evidence.body[i].srcAddress + "、";
                }

                worksheet.Cells[6, 5].Value += "\n";
            }

            //填写风险资产外网
            foreach (ResponseData.RAData.RAInfo raInfo in NetHelper.exNetList) {

                worksheet.Cells[7, 5].Value += raInfo.tags + ":" + raInfo.assetName + "\n";

            }

            //填写风险资产内网
            foreach (ResponseData.RAData.RAInfo raInfo in NetHelper.inNetList) {

                worksheet.Cells[8, 5].Value += raInfo.tags + ":" + raInfo.assetName + "\n";

            }

            //填写风险资产工控网
            foreach (ResponseData.RAData.RAInfo raInfo in NetHelper.icNetList) {

                worksheet.Cells[7, 5].Value += raInfo.tags + ":" + raInfo.assetName + "\n";

            }

            //填写风险资产其他
            foreach (ResponseData.RAData.RAInfo raInfo in NetHelper.otherNetList) {

                worksheet.Cells[7, 5].Value += raInfo.tags + ":" + raInfo.assetName + "\n";

            }

            package.SaveAs(newFile);
            package.Dispose();
        }

        //public static void Test() {
        //    FileInfo existingFile = new FileInfo(rzoaSheet);
        //    Console.WriteLine(Application.StartupPath);
        //}
        public static void Check() {
            try {
                foreach (String key in excelDic.Keys) {
                    FileInfo existingFile = new FileInfo(excelDic[key]);
                    ExcelPackage package = new ExcelPackage(existingFile);
                    //ExcelWorksheet worksheet = package.Workbook.Worksheets[1];//选定 指定页 默认第一个Sheet
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                    worksheet.Cells[1, 11].Value = "测试";
                    if (key == "erp") {
                        for (int i = 2; i <= worksheet.Dimension.Rows; i++) {
                            //worksheet.Cells[i, 11].Value = 
                            //    NetHelper.ErpRequset(worksheet.Cells[i, 9].Text,
                            //    worksheet.Cells[i, 10].Text) ? "成功" : "失败";
                        }
                    }
                    else {
                        for (int i = 2; i <= worksheet.Dimension.Rows; i++) {
                            //worksheet.Cells[i, 11].Value =
                            //    NetHelper.RzoaRequset(worksheet.Cells[i, 9].Text,
                            //    worksheet.Cells[i, 10].Text) ? "成功" : "失败";
                        }
                    }
                    package.Save();
                    package.Dispose();
                }
            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
                MessageBox.Show(e.Message);
            }
        }
    }
}
