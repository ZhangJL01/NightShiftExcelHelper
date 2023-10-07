using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using UnityEngine;
//using System.Windows.Forms;

namespace NightShiftExcelHelper {
    class ExcelHelper {
        private static readonly string time = DateTime.Now.ToString("yyyyMMdd");

        //模板路径
        public static readonly string nsSheet =
            Application.streamingAssetsPath + "/TemFiles/信息网络安全监控报表.xlsx";

        public static readonly string wpSheet =
            Application.streamingAssetsPath + "/TemFiles/OA、ERP弱口令信息.xlsx";
        
        private static readonly string wpSheetNew =
            Application.streamingAssetsPath + "/TemFiles/OA、ERP弱口令信息_NEW.xlsx";

        private static readonly string wpDoc =
            Application.streamingAssetsPath + "/TemFiles/弱口令截图.docx";

        private static readonly string wpSheetDT =
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/OA、ERP弱口令信息_" + time + ".xlsx";

        private static readonly string nsSheetDT =
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/信息网络安全监控报表_" + time + ".xlsx";

        private static readonly string wpDocDT =
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/弱口令截图_" + time + ".docx";

        public static Dictionary<String, String> excelDic = new Dictionary<string, string>();

        public static string FillWPSheet() {
            try {
                // 添加这行代码后不会报ExcelPackage错误
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                FileInfo fileInfo;
                ExcelPackage package;

                if (DateTime.Now.Day.Equals(1)) {
                    fileInfo = new FileInfo(wpSheetNew);
                    package = new ExcelPackage(fileInfo);

                    package.SaveAs(wpSheet);
                    package.Dispose();
                }

                fileInfo = new FileInfo(wpSheet);
                package = new ExcelPackage(fileInfo);

                WPSheetHandle(package.Workbook.Worksheets.First(), NetHelper.rzoaList);
                WPSheetHandle(package.Workbook.Worksheets.Last(), NetHelper.erpList);

                package.Save();
                fileInfo = new FileInfo(wpSheetDT);
                package.SaveAs(fileInfo);
                package.Dispose();
                File.Copy(wpDoc, wpDocDT);

                return null;
            } catch (Exception e) {
                return e.Message;
            }
        }

        private static void WPSheetHandle(ExcelWorksheet worksheet, List<ResponseData.WPData.WPInfo> list) {
            int curRow = worksheet.Dimension.End.Row + 1;
            if (list.Count == 0) {
                worksheet.Cells[curRow, 1].Value = "空";
                return;
            }
            worksheet.Cells["A" + curRow + ":A" + (curRow + list.Count - 1)].Merge = true;
            worksheet.Cells[curRow, 1].Value = DateTime.Now.ToString("MM.dd");
            list.ForEach(item => {
                worksheet.Cells[curRow, 2].Value = item.srcAddress;
                worksheet.Cells[curRow, 3].Value = item.destAddress;
                worksheet.Cells[curRow, 4].Value = item.alarmName;
                worksheet.Cells[curRow, 5].Value = item.subCategory;               
                worksheet.Cells[curRow, 6].Value = item.destPort;
                worksheet.Cells[curRow, 7].Value = item.threatSeverity;
                worksheet.Cells[curRow, 8].Value = item.alarmResults;
                worksheet.Cells[curRow, 9].Value = item.collectorReceiptTime;
                worksheet.Cells[curRow, 10].Value = item.srcUserName;
                worksheet.Cells[curRow, 11].Value = item.passwd;
                curRow++;
            });
            
        }

        public static void FillNSSheet() {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            FileInfo fileInfo = new FileInfo(nsSheet);

            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(DateTime.Now.ToString("MM.dd") + "信息网络安全监控报表", package.Workbook.Worksheets.Last());

            //数量
            worksheet.Cells[4, 4].Value = NetHelper.highSIList.Count;
            worksheet.Cells[5, 4].Value = NetHelper.mediumSIList.Count;
            worksheet.Cells[6, 4].Value = NetHelper.lowSIList.Count;
            worksheet.Cells[7, 4].Value = NetHelper.exNetList.Count;
            worksheet.Cells[8, 4].Value = NetHelper.inNetList.Count;
            worksheet.Cells[9, 4].Value = NetHelper.icNetList.Count;
            worksheet.Cells[10, 4].Value = NetHelper.otherNetList.Count;

            worksheet.Cells[1, 1].Value = "信息网络安全监控报表(" +
                DateTime.Now.Month + "月" + (DateTime.Now.Day - 1) + "日-" + DateTime.Now.Month + "月" + DateTime.Now.Day + "日)";
            worksheet.Cells[2, 1].Value = "监控人员:        填表人员:";

            for (int i = 4; i < 16; i++) {
                worksheet.Cells[i, 5].Value = "";
                worksheet.Cells[i, 6].Value = "";
            }
        
            //填写高危事件
            foreach (ResponseData.SIData.SIInfo siInfo in NetHelper.highSIList) {

                worksheet.Cells[4, 5].Value += siInfo.incidentName + ":";
                for (int i = 0; i < siInfo.evidence.body.Count; i++) {
                    worksheet.Cells[4, 5].Value += 
                        siInfo.evidence.body[i].srcAddress + siInfo.evidence.body[i].destAddress;
                    worksheet.Cells[4, 5].Value +=
                        i == siInfo.evidence.body.Count - 1 ? "" : "、";
                }

                worksheet.Cells[4, 5].Value += "\n";
            }

            //填写中危事件
            foreach (ResponseData.SIData.SIInfo siInfo in NetHelper.mediumSIList) {

                worksheet.Cells[5, 5].Value += siInfo.incidentName + ":";
                for (int i = 0; i < siInfo.evidence.body.Count; i++) {
                    worksheet.Cells[5, 5].Value += 
                        siInfo.evidence.body[i].srcAddress + siInfo.evidence.body[i].destAddress;
                    worksheet.Cells[5, 5].Value +=
                        i == siInfo.evidence.body.Count - 1 ? "" : "、";
                }

                worksheet.Cells[5, 5].Value += "\n";
            }

            //填写低危事件
            foreach (ResponseData.SIData.SIInfo siInfo in NetHelper.lowSIList) {

                worksheet.Cells[6, 5].Value += siInfo.incidentName + ":";
                for (int i = 0; i < siInfo.evidence.body.Count; i++) {
                    worksheet.Cells[6, 5].Value += 
                        siInfo.evidence.body[i].srcAddress + siInfo.evidence.body[i].destAddress;
                    worksheet.Cells[6, 5].Value +=
                        i == siInfo.evidence.body.Count - 1 ? "" : "、";
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

                worksheet.Cells[9, 5].Value += raInfo.tags + ":" + raInfo.assetName + "\n";

            }

            //填写风险资产其他
            foreach (ResponseData.RAData.RAInfo raInfo in NetHelper.otherNetList) {

                worksheet.Cells[10, 5].Value += raInfo.tags + ":" + raInfo.assetName + "\n";

            }

            package.Workbook.Worksheets.MoveToStart(worksheet.Name);
            worksheet.View.SetTabSelected();

            package.Save();

            if (DateTime.Now.Day.Equals(1)) {
                while(package.Workbook.Worksheets.Count > 1) {
                    package.Workbook.Worksheets.Delete(1);
                }
                package.Save();
            }
            fileInfo = new FileInfo(nsSheetDT);
            package.SaveAs(fileInfo);
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
                //MessageBox.Show(e.Message);
            }
        }
    }
}
