using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;

namespace ExcelTool.Services
{
    public class ExcelService
    {
        /// <summary>
        /// 默认程序excel路径
        /// </summary>
        private static string  _excelInPath= Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ConfigurationManager.AppSettings["ExcelPath"]);
        /// <summary>
        /// 输出路径
        /// </summary>
        private static string _excelOutPath= Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ConfigurationManager.AppSettings["ExcelOutput"]);
        
        public ExcelService() { 

        }
        /// <summary>
        /// 合并excel
        /// </summary>
        public static void Merge()
        {
            //创建一个新的workbook对象
            Workbook newbook = new Workbook();
            newbook.Version = ExcelVersion.Version2013;
            // 删除文档中的工作表（新创建的文档默认包含3张工作表）
             newbook.Worksheets.Clear();
            //创建一个临时的workbook，用于加载需要合并的Excel文档
            Workbook tempbook = new Workbook();
            var files= FileService.GetFiles(_excelInPath);
            
            //过滤
            foreach (var excle in files)
            {
                //加载数据
                tempbook.LoadFromFile(excle);
                //把sheet 复制到newbook 里
                newbook.Worksheets.AddCopy(tempbook.Worksheets[0]);
                MergeSheets(newbook);
                tempbook = null;
            }
            newbook.SaveToFile($@"{_excelOutPath}/mergedexcel{Guid.NewGuid()}.xlsx");

        }
        /// <summary>
        /// 合并csv文件
        /// </summary>
        public static void MergeCsv()
        {
            var files = FileService.GetFiles(_excelInPath);
            //
            var csvfiles = files?.Where(f => f.EndsWith(".csv")).ToList();
            var mergedfileName = $@"{_excelOutPath}/mergedexcel{Guid.NewGuid()}.xlsx";
            // 合并csv
            var sw = File.CreateText(mergedfileName);
            if (csvfiles != null && csvfiles.Count() > 0)
            {
                int i = 0;
                foreach (var file in csvfiles)
                {
                    // add to mergedfile 
                    int j = 0;
                    using (var stream = new StreamReader(file))
                    {
                        var line = stream.ReadLine();
                        while (!string.IsNullOrEmpty(line))
                        {
                            if (i > 0 && j == 0)
                            {                             
                            }
                            else
                            {
                                sw.WriteLine(line);
                            }
                            line = stream.ReadLine();
                            i++;
                        }                         
                    }
                    j++;
                }
                sw.Flush();
                sw.Close();
            }
        }
        public static Workbook MergeSheets(Workbook workbook)
        {
            if (workbook.Worksheets.Count() > 1)
            {
                var sheet1 = workbook.Worksheets[0];
                var sheet2 = workbook.Worksheets[1];
                sheet2.AllocatedRange.Copy(sheet1[sheet1.LastRow + 1,1]);
                sheet2.Remove();
            }
            return workbook;
        }

    }
}
