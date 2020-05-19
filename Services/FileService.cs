using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelTool.Services
{
    class FileService
    {

        /// <summary>
        /// 获取文件
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static string[] GetFiles(string path)
        {
            if (!Directory.Exists(path))
                return null;
            return Directory.GetFiles(path);
        }
    }
}
