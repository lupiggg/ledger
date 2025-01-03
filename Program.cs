using System;
using System.IO;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace ledgerByLupig
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // 指定文件夹路径
            string folderPath = @"D:\USERS\Administrator\Desktop\tz";
            DirectoryInfo folder = new DirectoryInfo(folderPath);

            // 检查文件夹是否存在
            if (!folder.Exists)
            {
                Console.WriteLine("指定的桌面路径不是一个有效的文件夹,请重新检查。");
                Console.WriteLine("PS:需要把要处理的台账放到【D:/USERS/Administrator/Desktop/tz】路径后再执行哦");
                Console.WriteLine("可按任意键退出...");
                // 等待用户按下任意键
                Console.ReadKey();
                return;
            }

            // 创建一个新的Excel工作簿
            using (var workbook = new XLWorkbook())
            {
                IXLWorksheet sheet = workbook.Worksheets.Add("File List");

                // 创建第一行表头
                sheet.Cell(1, 1).Value = "序号";
                sheet.Cell(1, 2).Value = "合同签订日期";
                sheet.Cell(1, 3).Value = "合同相对方";
                sheet.Cell(1, 4).Value = "合同业务内容";
                sheet.Cell(1, 5).Value = "开票";

                // 获取文件夹中的所有文件
                FileInfo[] fileList = folder.GetFiles();
                int rowNum = 2; // 从第二行开始写入数据
                foreach (FileInfo file in fileList)
                {
                    // 去掉开头的"322."和结尾的".pdf"
                    string trimmedString = file.Name.Substring(4, file.Name.Length - 8);

                    // 使用连字符分割字符串
                    string[] splitStrings = trimmedString.Split(new[] { '-' }, 3, StringSplitOptions.None);

                    // 找到第一个和最后一个连字符的位置
                    int firstDashIndex = trimmedString.IndexOf('-');
                    int lastDashIndex = trimmedString.LastIndexOf('-');

                    // 分割字符串
                    string companyName = trimmedString.Substring(0, firstDashIndex).Trim();
                    string salesDetails = trimmedString.Substring(firstDashIndex + 1, lastDashIndex - firstDashIndex - 1).Trim();
                    string invoiceStatus = trimmedString.Substring(lastDashIndex + 1).Trim();

                    // 定义正则表达式
                    string pattern = "(.*?)\\（(.*?)\\）";
                    Regex r = new Regex(pattern);
                    Match matcher = r.Match(salesDetails);

                    if (matcher.Success)
                    {
                        salesDetails = matcher.Groups[1].Value; // 提取第一个匹配组
                        string date = matcher.Groups[2].Value; // 提取第二个匹配组

                        IXLRow row = sheet.Row(rowNum++);
                        row.Cell(1).Value = file.Name.Split('.')[0];
                        row.Cell(2).Value = date;
                        row.Cell(3).Value = companyName;
                        row.Cell(4).Value = salesDetails;
                        row.Cell(5).Value = invoiceStatus;
                    }
                    else
                    {
                        Console.WriteLine("未找到匹配项");
                        Console.WriteLine("可按任意键退出...");
                        // 等待用户按下任意键
                        Console.ReadKey();
                    }
                }

                // 写入Excel文件
                workbook.SaveAs("D:\\USERS\\Administrator\\Desktop\\台账（提取后）.xlsx");
                Console.WriteLine("全部台账信息已成功写入Excel,文件路径在：【D:/USERS/Administrator/Desktop】");
                Console.WriteLine("可按任意键退出...");
                // 等待用户按下任意键
                Console.ReadKey();
            }
        }
    }
}
