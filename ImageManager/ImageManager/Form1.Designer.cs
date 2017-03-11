using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel; //使用别名
using System.Reflection;
using System;
using System.Security.Cryptography;
using System.Text;
using System.Linq;
using System.Data.SqlClient;
using System.Data;
namespace ImageManager
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// 获取用户选择的文件夹路径，返回该路径
        /// </summary>
        public string GetImagPath()
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择目标文件夹";
            dialog.ShowDialog();
            string path = dialog.SelectedPath;
            textBoxOfResultForCreateExcel.Text = path + "\n";
            return path;
        }

        /// <summary>
        /// 获取目标目录下的所有jpg文件，并将文件路径以数组形式返回
        /// </summary>
        /// <returns>所有的文件路径，string数组</returns>
        public string[] GetFilePath()
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择目标文件夹";
            dialog.ShowDialog();
            string path = dialog.SelectedPath;
            string[] imagePath = Directory.GetFiles(path, "*.jpg", SearchOption.AllDirectories);
            return imagePath;
        }

        /// <summary>
        /// 获取目标目录下的所有excel文件，并将文件路径以数组形式返回
        /// </summary>
        /// <returns>所有文件路径，string数组</returns>
        public string[] GetExcelsPath()
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择目标文件夹";
            dialog.ShowDialog();
            string path = dialog.SelectedPath;
            string[] excelPath = Directory.GetFiles(path, "*.xlsx", SearchOption.AllDirectories).Union(Directory.GetFiles(path, "*.csv", SearchOption.AllDirectories)).ToArray<string>();
            return excelPath;
        }

        /// <summary>
        /// 根据图片路径获取学号并编码
        /// </summary>
        /// <param name="imagePath">图片路径</param>
        /// <returns>编码结果</returns>
        public string GetCodes(string imagePath)
        {
            int lastIndex = imagePath.LastIndexOf(@"\");
            string temp = imagePath.Remove(0, lastIndex + 1);
            string studentId = null;
            if (temp.Contains("（"))
            {
                studentId = temp.Split(new string[] { "（" }, 5, StringSplitOptions.RemoveEmptyEntries)[0];
            }
            else if (temp.Contains("("))
            {
                studentId = temp.Split(new string[] { "(" }, 5, StringSplitOptions.RemoveEmptyEntries)[0];
            }
            else
            {
                System.Windows.Forms.Application.Exit();
            }
            //加密过程就不能暴露再github上啦,嘿嘿嘿
            string md5Code = "";
            return md5Code;
        }

        /// <summary>
        /// </summary>
        /// <param name="imagePath">源文件路径</param>
        /// <param name="code">编码结果</param>
        public void MoveImage(string imagePath, string storagePath, string code)
        {
            string path1 = code[1].ToString();
            string path2 = code[3].ToString();;
            string finalPath = storagePath + @"\" + path1 + @"\" + path2;
            if (Directory.Exists(finalPath))
            {
                try
                {
                    File.Copy(imagePath, finalPath + @"\" + code + @".jpg", false);
                    textBoxOfResultForMoveImage.AppendText("移动" + imagePath + "成功！\n");
                }
                catch (IOException e)
                {
                    File.Copy(imagePath, finalPath + @"\" + code + "2" + @".jpg", false);
                    textBoxOfResultForMoveImage.AppendText("移动" + imagePath + "成功！\n");
                }


            }
            else
            {
                Directory.CreateDirectory(finalPath);
                try
                {
                    File.Copy(imagePath, finalPath + @"\" + code + @".jpg", false);
                    textBoxOfResultForMoveImage.AppendText("移动" + imagePath + "成功！\n");
                }
                catch (IOException e)
                {
                    File.Copy(imagePath, finalPath + @"\" + code + "2" + @".jpg", false);
                    textBoxOfResultForMoveImage.AppendText("移动" + imagePath + "成功！\n");
                }
            }
        }

        /// <summary>
        /// 根据Excel表格路径将该表格添加到总表中
        /// </summary>
        /// <param name="strFileName">原表格路径</param>
        public void OpenExcel(string strFileName)
        {
            object missing = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.Application();//lauch excel application  

            if (excel == null)
            {
                textBoxOfResultForCreateExcel.AppendText("Can't access excel" + "  " + strFileName);
            }
            else
            {
                excel.Visible = false; excel.UserControl = true;
                // 以只读的形式打开EXCEL文件  
                Excel.Workbook wb = excel.Application.Workbooks.Open(strFileName, missing, true, missing, missing, missing,
                 missing, missing, missing, true, missing, missing, missing, missing, missing);
                //取得第一个工作薄  
                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);
                //取得总记录行数    (包括标题列)  
                int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数  
                //int columnsint = mySheet.UsedRange.Cells.Columns.Count;//得到列数  
                //取得数据范围区域   (不包括标题列)    
                Excel.Range rng1 = ws.Cells.get_Range("A2", "A" + rowsint);
                Excel.Range rng2 = ws.Cells.get_Range("B2", "B" + rowsint);
                Excel.Range rng3 = ws.Cells.get_Range("C2", "C" + rowsint);
                Excel.Range rng4 = ws.Cells.get_Range("D2", "D" + rowsint);
                object[,] arry1 = (object[,])rng1.Value2;   //get range's value ,学号 
                object[,] arry2 = (object[,])rng2.Value2;   //专业
                object[,] arry3 = (object[,])rng3.Value2;   //get range's value  姓名
                object[,] arry4 = (object[,])rng4.Value2;   //性别
                //将arry1赋给一个数组  
                string classId = arry1[2, 1].ToString().Remove(arry1[2, 1].ToString().Length - 2);
                //新建Excelapp
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = true;
                //设置不显示确认修改提示
                excelApp.DisplayAlerts = false;
                //得到workbook对象，打开总表
                Excel.Workbook excleBook = excelApp.Workbooks.Open(@"D:\总表.xlsx", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                //指定要操作的sheet
                Excel._Worksheet excleSheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(1);
                //获取总记录行数（包括标题列）
                int rowsCount = excleSheet.UsedRange.Cells.Rows.Count;
                //将分表的记录写入到总表中去
                for (int i = 1; i <= arry1.Length; i++)
                {
                    if (arry1[i, 1] == null)
                    {
                        excelApp.Cells[rowsCount + i, 1] = "";
                    }
                    else
                    {
                        excelApp.Cells[rowsCount + i, 1] = arry1[i, 1].ToString();
                    }
                    if (arry2[i, 1] == null)
                    {
                        excelApp.Cells[rowsCount + i, 2] = "";
                    }
                    else
                    {
                        excelApp.Cells[rowsCount + i, 2] = arry2[i, 1].ToString();
                    }
                    if (arry3[i, 1] == null)
                    {
                        excelApp.Cells[rowsCount + i, 3] = "";
                    }
                    else
                    {
                        if (arry3[i, 1].ToString().EndsWith("0001"))
                        {
                            excelApp.Cells[rowsCount + i, 3] = arry3[i, 1].ToString().Remove(arry3[i, 1].ToString().Length - 4);
                        }
                        else
                        {
                            excelApp.Cells[rowsCount + i, 3] = arry3[i, 1].ToString();
                        }

                    }
                    if (arry4[i, 1] == null)
                    {
                        excelApp.Cells[rowsCount + i, 4] = "";
                    }
                    else
                    {
                        excelApp.Cells[rowsCount + i, 4] = arry4[i, 1].ToString();
                    }

                    excelApp.Cells[rowsCount + i, 5] = classId;

                }
                //for (int i = 1; i <= rowsint - 1; i++)  
                //    for (int i = 1; i <= rowsint - 2; i++)
                //    {

                //        arry[i - 1, 0] = arry1[i, 1].ToString();

                //        arry[i - 1, 1] = arry2[i, 1].ToString();

                //        arry[i - 1, 2] = arry3[i, 1].ToString();

                //        arry[i - 1, 3] = arry4[i, 1].ToString();
                //    }
                //    string a = "";
                //    for (int i = 0; i <= rowsint - 3; i++)
                //    {
                //        a += arry[i, 0] + "|" + arry[i, 1] + "|" + arry[i, 2] + "|" + arry[i, 3] + "\n";

                //    }
                //    this.label1.Text = a;
                //关闭打开的表

                wb.Close(false, Missing.Value, Missing.Value);
                excleBook.Close(true, Missing.Value, Missing.Value);
                //把excel，workbook，worksheet设置为null，防止内存泄漏
                wb = null;
                excleBook = null;
                ws = null;
                excleBook = null;
                //程序退出  
                excel.Quit();
                excelApp.Quit();
                excel = null;
                excelApp = null;
                GC.Collect();
            }
        }

        /// <summary>
        /// 创建excel表格
        /// </summary>
        /// <param name="path">目标文件夹的路径</param>
        public void CreateExcel(string path)
        {
            string[] imagPath = Directory.GetFiles(path, "*.jpg", SearchOption.AllDirectories);
            string[] imagName = new string[imagPath.Length];
            string classId = path;
            List<string> studentsId = new List<string>();
            List<string> speciality = new List<string>();
            List<string> StudentsName = new List<string>();
            List<string> studentGender = new List<string>();
            //拿到目标文件夹下的文件名并且分割存储
            for (int i = 0; i < imagPath.Length; i++)
            {
                imagName[i] = Path.GetFileNameWithoutExtension(imagPath[i]);
                if (imagName[i].EndsWith("0003") || imagName[i].EndsWith("0004"))
                {
                    textBoxOfResultForCreateExcel.AppendText("出现未处理的" + imagName[i] + "\n");
                }
                if (imagName[i].Contains("0001") && imagName[i].EndsWith("0001"))
                {
                    //探查是否存在0002图片，如果存在则将该人写入excel，否则不写入
                    if( !File.Exists(imagPath[i].Remove(imagPath[i].Length - 8, 8) + "0002.jpg"))
                    {
                        continue;
                    }
                    string[] fields = new string[5];
                    //版本1下图片名称处理代码块
                    //if(imagName[i].Contains("(") || imagName[i].Contains(")"))
                    //{
                    //    fields = imagName[i].Split(new string[] { "(", ")" }, 3, System.StringSplitOptions.None);
                    //}
                    //if(imagName[i].Contains("（") || imagName[i].Contains("）"))
                    //{
                    //    fields = imagName[i].Split(new string[] { "（", "）" }, 3, System.StringSplitOptions.None);
                    //}
                    //if (fields == null)
                    //{
                    //    fields = new string[3] { " ", " ", " " };
                    //}
                    //版本2下图片名称处理代码块
                    if (imagName[i].Contains("(") || imagName[i].Contains(")"))
                    {
                        fields = imagName[i].Split(new string[] { "(", ")" }, 5, System.StringSplitOptions.None);
                    }
                    if (imagName[i].Contains("（") || imagName[i].Contains("）"))
                    {
                        fields = imagName[i].Split(new string[] { "（", "）" }, 5, System.StringSplitOptions.None);
                    }
                    if (fields == null)
                    {
                        fields = new string[4] { " ", " ", " ", " " };
                    }
                    studentsId.Add(fields[0]);
                    speciality.Add(fields[1]);
                    if (fields.Length <= 3)
                    {
                        fields[2] = " ";
                        StudentsName.Add(fields[2]);
                        studentGender.Add(" ");
                    }
                    else
                    {
                        StudentsName.Add(fields[2]);
                        studentGender.Add(fields[3]);
                    }


                }

            }
            //Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();
            //_Workbook book = xls.Workbooks.Add(Missing.Value);
            //_Worksheet sheet;
            //xls.Visible = false;
            //xls.DisplayAlerts = true;
            //for(int i = 1; i <= 3; i++ )
            //{
            //    //增加一个sheet
            //    sheet= (_Worksheet)book.Worksheets.Add(Missing.Value, book.Worksheets[book.Sheets.Count], 1, Missing.Value);
            //    sheet.Name = "学号";
            //    for (int row = 1; row <= studentsId.Count; row++)//循环设置每个单元格的值
            //    {
            //        for (int offset = 1; offset < 10; offset++)
            //            sheet.Cells[row, offset] = "( " + row.ToString() + "," + offset.ToString() + " )";
            //    }
            //}


            int nMax = studentsId.Count;
            //int nMin = 1;

            int rowCount = nMax;//总行数  


            //版本1下的Excel处理代码
            //const int columnCount = 3;//总列数  


            //版本2下的Excel处理代码
            const int columnCount = 4;//总列数


            //创建Excel对象  

            Excel.Application excelApp = new Excel.Application();



            //新建工作簿  

            Excel.Workbook workBook = excelApp.Workbooks.Add(true);



            //新建工作表  

            Excel.Worksheet worksheet = workBook.ActiveSheet as Excel.Worksheet;



            ////设置标题  

            //Excel.Range titleRange = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, columnCount]);//选取单元格  

            //titleRange.Merge(true);//合并单元格  

            //titleRange.Value2 = strTitle; //设置单元格内文本  

            //titleRange.Font.Name = "宋体";//设置字体  

            //titleRange.Font.Size = 18;//字体大小  

            //titleRange.Font.Bold = false;//加粗显示  

            //titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//水平居中  

            //titleRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;//垂直居中  

            //titleRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;//设置边框  

            //titleRange.Borders.Weight = Excel.XlBorderWeight.xlThin;//边框常规粗细  



            //设置表头  

            //版本1 的代码控制Excel列
            //string[] strHead = new string[columnCount] { "学号", "专业", "姓名" };

            //int[] columnWidth = new int[3] { 8, 16, 8};

            //版本2的代码控制Excel列
            string[] strHead = new string[columnCount] { "学号", "专业", "姓名", "性别" };

            int[] columnWidth = new int[4] { 16, 16, 16, 16 };
            for (int i = 0; i < columnCount; i++)
            {

                //Excel.Range headRange = worksheet.Cells[2, i + 1] as Excel.Range;//获取表头单元格  

                Excel.Range headRange = worksheet.Cells[1, i + 1] as Excel.Range;//获取表头单元格,不用标题则从1开始  

                headRange.Value2 = strHead[i];//设置单元格文本  

                headRange.Font.Name = "宋体";//设置字体  

                headRange.Font.Size = 12;//字体大小  

                headRange.Font.Bold = false;//加粗显示  

                headRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//水平居中  

                headRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;//垂直居中  

                headRange.ColumnWidth = columnWidth[i];//设置列宽  

                //  headRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;//设置边框  

                // headRange.Borders.Weight = Excel.XlBorderWeight.xlThin;//边框常规粗细  

            }



            //设置每列格式  

            //for (int i = 0; i < columnCount; i++)
            //{

            //    //Excel.Range contentRange = worksheet.get_Range(worksheet.Cells[3, i + 1], worksheet.Cells[rowCount - 1 + 3, i + 1]);  

            //    //Excel.Range contentRange = worksheet.get_Range(worksheet.Cells[3, i + 1], worksheet.Cells[rowCount - 1 + 3, i + 1]);//不用标题则从第二行开始  


            //    contentRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//水平居中  

            //    contentRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;//垂直居中  

            //    //contentRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;//设置边框  

            //    // contentRange.Borders.Weight = Excel.XlBorderWeight.xlThin;//边框常规粗细  

            //    contentRange.WrapText = true;//自动换行  

            //    contentRange.NumberFormatLocal = "@";//文本格式  

            //}



            //填充数据  

            for (int i = 0; i < nMax; i++)
            {

                int k = i;

                //excelApp.Cells[k + 3, 1] = string.Format("{0}", k + 1);  

                //excelApp.Cells[k + 3, 2] = string.Format("{0}-{1}", i - 0.5, i + 0.5);  

                //excelApp.Cells[k + 3, 3] = string.Format("{0}", k + 3);  

                //excelApp.Cells[k + 3, 4] = string.Format("{0}", k + 4);  

                excelApp.Cells[k + 2, 1] = studentsId[k];

                excelApp.Cells[k + 2, 2] = speciality[k];

                excelApp.Cells[k + 2, 3] = StudentsName[k];

                excelApp.Cells[k + 2, 4] = studentGender[k];

            }



            //设置Excel可见  

            excelApp.Visible = true;

            //保存文件
            workBook.SaveAs(classId, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            //关闭打开的表
            workBook.Close(false, Missing.Value, Missing.Value);
            //把excel，workbook，worksheet设置为null，防止内存泄漏
            workBook = null;
            worksheet = null;
            //程序退出  
            excelApp.Quit();
            excelApp = null;
            GC.Collect();
        }

        /// <summary>
        /// 将excel中的数据导入到数据库中
        /// </summary>
        /// <param name="path">excel文件的路径</param>
        /// <returns></returns>
        public bool LoadExcelData(string path)
        {
           object missing = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.Application();//lauch excel application  

            if (excel == null)
            {
                textBoxOfResultForCreateExcel.AppendText("Can't access excel" + "  " + path);
                return false;
            }
            else
            {
                excel.Visible = false; excel.UserControl = true;
                // 以只读的形式打开EXCEL文件  
                Excel.Workbook wb = excel.Application.Workbooks.Open(path, missing, true, missing, missing, missing,
                 missing, missing, missing, true, missing, missing, missing, missing, missing);
                //取得第一个工作薄  
                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);
                //取得总记录行数    (包括标题列)  
                int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数  
                //int columnsint = mySheet.UsedRange.Cells.Columns.Count;//得到列数  
                //取得数据范围区域   (不包括标题列)    
                Excel.Range rng1 = ws.Cells.get_Range("A2", "A" + rowsint);  // 序号
                Excel.Range rng2 = ws.Cells.get_Range("B2", "B" + rowsint);  //班级
                Excel.Range rng3 = ws.Cells.get_Range("C2", "C" + rowsint); //学号
                Excel.Range rng4 = ws.Cells.get_Range("D2", "D" + rowsint); //姓名
                Excel.Range rng5 = ws.Cells.get_Range("E2", "E" + rowsint); //专业
                Excel.Range rng6 = ws.Cells.get_Range("F2", "F" + rowsint); //性别
                Excel.Range rng7 = ws.Cells.get_Range("G2", "G" + rowsint); //学院
                Excel.Range rng8 = ws.Cells.get_Range("H2", "H" + rowsint); //入学年份
                Excel.Range rng9 = ws.Cells.get_Range("I2", "I" + rowsint); //年级
                object[,] serialNumber = (object[,])rng1.Value2;   //get range's value ,序号
                object[,] classNumber = (object[,])rng2.Value2;   //班级
                object[,] studentID = (object[,])rng3.Value2;   //get range's value  学号
                object[,] name = (object[,])rng4.Value2;   //姓名
                object[,] major = (object[,])rng5.Value2;  //专业
                object[,] gender = (object[,])rng6.Value2;  //性别
                object[,] institute = (object[,])rng7.Value2;  //学院
                object[,] entranceTime = (object[,])rng8.Value2;  //入学年份
                object[,] grade = (object[,])rng9.Value2;  //年级
                for (int i = 1; i <= rowsint - 1; i++ )
                {
                    if (gender[i, 1] == null) gender[i, 1] = "";
                    if (gender[i, 1].ToString() == "男") gender[i, 1] = "1";
                    if (gender[i, 1].ToString() == "女") gender[i, 1] = "0";
                    string sql = "INSERT INTO StudentInfo (serialNumber, classNumber, studentID, name, major, gender, institute, entranceTime, grade) VALUES (@serialNumber, @classNumber, @studentID, @name, @major, @gender, @institute, @entranceTime, @grade)";
                    int isSucceed = SqlHelper.ExecuteNonQuery(sql, 
                        new SqlParameter("@serialNumber", serialNumber[i, 1]),
                        new SqlParameter("@classNumber", classNumber[i, 1] == null ? "" : classNumber[i, 1]),
                        new SqlParameter("@studentID", studentID[i, 1] == null ? "" : studentID[i, 1]),
                        new SqlParameter("@name", name[i, 1] == null ? "" : name[i, 1]),
                        new SqlParameter("@major", major[i, 1] == null ? "" : major[i, 1]),
                        new SqlParameter("@gender", gender[i, 1]),
                        new SqlParameter("@institute", institute[i, 1] == null ? "" : institute[i, 1]),
                        new SqlParameter("@entranceTime", entranceTime[i, 1] == null ? "" : entranceTime[i, 1]),
                        new SqlParameter("@grade", grade[i, 1] == null ? "" : grade[i, 1]));
                    if (isSucceed == 0)
                    {
                        textBoxOfResultForCreateExcel.AppendText(string.Format("序号 = {0}，学号 = {1}， 姓名 = {2}插入失败！\n", serialNumber[i, 1], studentID[i, 1], name[i, 1]));
                    }
                    else
                    {
                        textBoxOfResultForCreateExcel.AppendText(string.Format("序号 = {0}，学号 = {1}， 姓名 = {2}插入成功！\n", serialNumber[i, 1], studentID[i, 1], name[i, 1]));
                    }
                }
                    
                return true;
            }
        }

        /// <summary>
        /// 审查图片信息是否匹配数据库中的信息
        /// </summary>
        /// <param name="path">图片文件夹路径</param>
        /// <returns></returns>
        public void CheckImageInfo(string path)
        {
            string[] imagePath = Directory.GetFiles(path, "*.jpg", SearchOption.AllDirectories);
            for(int i = 0; i <= imagePath.Length - 1; i++)
            {
                int lastIndex = imagePath[i].LastIndexOf(@"\");
                string strTemp = imagePath[i].Remove(0, lastIndex + 1);
                string[] info = strTemp.Split(new string[]{"(", ")", "（", "）"}, StringSplitOptions.RemoveEmptyEntries);
                //处理通达学院
                if(info[0].StartsWith("M"))
                {
                    info[0] = info[0].Replace("M", "");
                }
                if (info.Length < 5)
                {
                    textBoxOfResultForCreateExcel.AppendText("错误：" + imagePath[i] + "\n");
                    continue;
                }
                if(SQLCheck(info) == false)
                {
                    textBoxOfResultForCreateExcel.AppendText("错误：" + imagePath[i] + "\n");
                }
            }
            textBoxOfResultForCreateExcel.AppendText("审查完毕\n");
        }

        /// <summary>
        /// 查询数据库匹配info信息是否有误，只匹配学号、名字
        /// </summary>
        /// <param name="info">待匹配的字符串数组</param>
        /// <returns></returns>
        public bool SQLCheck(string[] info)
        {
            string studentID = info[0];
            string name = info[2];
            string gender = info[3];
            string sql = "SELECT * FROM StudentInfo WHERE studentID = @studentID AND name = @name";
            DataSet mDataSet = SqlHelper.ExecuteDataSet(sql, new SqlParameter("@studentID", studentID), new SqlParameter("@name", name));
            System.Data.DataTable mDataTable = mDataSet.Tables[0];
            if(mDataTable.Rows.Count == 0){
                return false;
            }
            if(mDataTable.Rows[0]["gender"].ToString() !="" && mDataTable.Rows[0]["gender"].ToString() != gender)
            {
                return false;
            }
            return true;
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonForFolderOfImages = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonForPathOfImages = new System.Windows.Forms.Button();
            this.buttonForFolderOfExcels = new System.Windows.Forms.Button();
            this.textBoxOfPathForCreateExcel = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxOfResultForCreateExcel = new System.Windows.Forms.TextBox();
            this.buttonForMoveImagesByFolder = new System.Windows.Forms.Button();
            this.buttonForMoveImagesByPath = new System.Windows.Forms.Button();
            this.textBoxOfPathForMoveImage = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxOfResultForMoveImage = new System.Windows.Forms.TextBox();
            this.buttonForSpareImages = new System.Windows.Forms.Button();
            this.buttonForLostImages = new System.Windows.Forms.Button();
            this.buttonForSQL = new System.Windows.Forms.Button();
            this.buttonForCheck = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonForFolderOfImages
            // 
            this.buttonForFolderOfImages.Location = new System.Drawing.Point(12, 43);
            this.buttonForFolderOfImages.Name = "buttonForFolderOfImages";
            this.buttonForFolderOfImages.Size = new System.Drawing.Size(155, 23);
            this.buttonForFolderOfImages.TabIndex = 0;
            this.buttonForFolderOfImages.Text = "选择文件夹（选取图片）";
            this.buttonForFolderOfImages.UseVisualStyleBackColor = true;
            this.buttonForFolderOfImages.Click += new System.EventHandler(this.buttonForFolderOfImages_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(241, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "制表区";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(241, 280);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "图片搬运区";
            // 
            // buttonForPathOfImages
            // 
            this.buttonForPathOfImages.Location = new System.Drawing.Point(12, 72);
            this.buttonForPathOfImages.Name = "buttonForPathOfImages";
            this.buttonForPathOfImages.Size = new System.Drawing.Size(167, 23);
            this.buttonForPathOfImages.TabIndex = 3;
            this.buttonForPathOfImages.Text = "选择下方路径（选取图片）";
            this.buttonForPathOfImages.UseVisualStyleBackColor = true;
            this.buttonForPathOfImages.Click += new System.EventHandler(this.buttonForPathOfImages_Click);
            // 
            // buttonForFolderOfExcels
            // 
            this.buttonForFolderOfExcels.Location = new System.Drawing.Point(12, 101);
            this.buttonForFolderOfExcels.Name = "buttonForFolderOfExcels";
            this.buttonForFolderOfExcels.Size = new System.Drawing.Size(155, 23);
            this.buttonForFolderOfExcels.TabIndex = 4;
            this.buttonForFolderOfExcels.Text = "选择文件夹（选取表格）";
            this.buttonForFolderOfExcels.UseVisualStyleBackColor = true;
            this.buttonForFolderOfExcels.Click += new System.EventHandler(this.buttonForFolderOfExcels_Click);
            // 
            // textBoxOfPathForCreateExcel
            // 
            this.textBoxOfPathForCreateExcel.Location = new System.Drawing.Point(12, 195);
            this.textBoxOfPathForCreateExcel.Multiline = true;
            this.textBoxOfPathForCreateExcel.Name = "textBoxOfPathForCreateExcel";
            this.textBoxOfPathForCreateExcel.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxOfPathForCreateExcel.Size = new System.Drawing.Size(239, 71);
            this.textBoxOfPathForCreateExcel.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(190, 48);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 12);
            this.label3.TabIndex = 7;
            this.label3.Text = "操作结果：";
            // 
            // textBoxOfResultForCreateExcel
            // 
            this.textBoxOfResultForCreateExcel.Location = new System.Drawing.Point(257, 63);
            this.textBoxOfResultForCreateExcel.Multiline = true;
            this.textBoxOfResultForCreateExcel.Name = "textBoxOfResultForCreateExcel";
            this.textBoxOfResultForCreateExcel.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxOfResultForCreateExcel.Size = new System.Drawing.Size(371, 203);
            this.textBoxOfResultForCreateExcel.TabIndex = 8;
            // 
            // buttonForMoveImagesByFolder
            // 
            this.buttonForMoveImagesByFolder.Location = new System.Drawing.Point(12, 310);
            this.buttonForMoveImagesByFolder.Name = "buttonForMoveImagesByFolder";
            this.buttonForMoveImagesByFolder.Size = new System.Drawing.Size(155, 23);
            this.buttonForMoveImagesByFolder.TabIndex = 9;
            this.buttonForMoveImagesByFolder.Text = "选择文件夹（选取图片）";
            this.buttonForMoveImagesByFolder.UseVisualStyleBackColor = true;
            this.buttonForMoveImagesByFolder.Click += new System.EventHandler(this.buttonForMoveImagesByFolder_Click);
            // 
            // buttonForMoveImagesByPath
            // 
            this.buttonForMoveImagesByPath.Location = new System.Drawing.Point(12, 339);
            this.buttonForMoveImagesByPath.Name = "buttonForMoveImagesByPath";
            this.buttonForMoveImagesByPath.Size = new System.Drawing.Size(167, 23);
            this.buttonForMoveImagesByPath.TabIndex = 10;
            this.buttonForMoveImagesByPath.Text = "选择下方路径（选取图片）";
            this.buttonForMoveImagesByPath.UseVisualStyleBackColor = true;
            this.buttonForMoveImagesByPath.Click += new System.EventHandler(this.buttonForMoveImagesByPath_Click);
            // 
            // textBoxOfPathForMoveImage
            // 
            this.textBoxOfPathForMoveImage.Location = new System.Drawing.Point(12, 368);
            this.textBoxOfPathForMoveImage.Multiline = true;
            this.textBoxOfPathForMoveImage.Name = "textBoxOfPathForMoveImage";
            this.textBoxOfPathForMoveImage.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxOfPathForMoveImage.Size = new System.Drawing.Size(243, 201);
            this.textBoxOfPathForMoveImage.TabIndex = 11;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(190, 315);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 12;
            this.label4.Text = "操作结果：";
            // 
            // textBoxOfResultForMoveImage
            // 
            this.textBoxOfResultForMoveImage.Location = new System.Drawing.Point(261, 327);
            this.textBoxOfResultForMoveImage.Multiline = true;
            this.textBoxOfResultForMoveImage.Name = "textBoxOfResultForMoveImage";
            this.textBoxOfResultForMoveImage.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxOfResultForMoveImage.Size = new System.Drawing.Size(378, 242);
            this.textBoxOfResultForMoveImage.TabIndex = 13;
            // 
            // buttonForSpareImages
            // 
            this.buttonForSpareImages.Location = new System.Drawing.Point(12, 131);
            this.buttonForSpareImages.Name = "buttonForSpareImages";
            this.buttonForSpareImages.Size = new System.Drawing.Size(109, 23);
            this.buttonForSpareImages.TabIndex = 14;
            this.buttonForSpareImages.Text = "查找多余图片";
            this.buttonForSpareImages.UseVisualStyleBackColor = true;
            this.buttonForSpareImages.Click += new System.EventHandler(this.buttonForSpareImages_Click);
            // 
            // buttonForLostImages
            // 
            this.buttonForLostImages.Location = new System.Drawing.Point(12, 160);
            this.buttonForLostImages.Name = "buttonForLostImages";
            this.buttonForLostImages.Size = new System.Drawing.Size(109, 23);
            this.buttonForLostImages.TabIndex = 15;
            this.buttonForLostImages.Text = "查找缺失图片";
            this.buttonForLostImages.UseVisualStyleBackColor = true;
            this.buttonForLostImages.Click += new System.EventHandler(this.buttonForLostImages_Click);
            // 
            // buttonForSQL
            // 
            this.buttonForSQL.Location = new System.Drawing.Point(127, 131);
            this.buttonForSQL.Name = "buttonForSQL";
            this.buttonForSQL.Size = new System.Drawing.Size(99, 23);
            this.buttonForSQL.TabIndex = 16;
            this.buttonForSQL.Text = "导入";
            this.buttonForSQL.UseVisualStyleBackColor = true;
            this.buttonForSQL.Click += new System.EventHandler(this.buttonForSQL_Click);
            // 
            // buttonForCheck
            // 
            this.buttonForCheck.Location = new System.Drawing.Point(127, 160);
            this.buttonForCheck.Name = "buttonForCheck";
            this.buttonForCheck.Size = new System.Drawing.Size(99, 23);
            this.buttonForCheck.TabIndex = 17;
            this.buttonForCheck.Text = "审查";
            this.buttonForCheck.UseVisualStyleBackColor = true;
            this.buttonForCheck.Click += new System.EventHandler(this.buttonForCheck_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(640, 627);
            this.Controls.Add(this.buttonForCheck);
            this.Controls.Add(this.buttonForSQL);
            this.Controls.Add(this.buttonForLostImages);
            this.Controls.Add(this.buttonForSpareImages);
            this.Controls.Add(this.textBoxOfResultForMoveImage);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBoxOfPathForMoveImage);
            this.Controls.Add(this.buttonForMoveImagesByPath);
            this.Controls.Add(this.buttonForMoveImagesByFolder);
            this.Controls.Add(this.textBoxOfResultForCreateExcel);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBoxOfPathForCreateExcel);
            this.Controls.Add(this.buttonForFolderOfExcels);
            this.Controls.Add(this.buttonForPathOfImages);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonForFolderOfImages);
            this.Name = "Form1";
            this.Text = "imageManager";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonForFolderOfImages;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button buttonForPathOfImages;
        private System.Windows.Forms.Button buttonForFolderOfExcels;
        private System.Windows.Forms.TextBox textBoxOfPathForCreateExcel;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxOfResultForCreateExcel;
        private System.Windows.Forms.Button buttonForMoveImagesByFolder;
        private System.Windows.Forms.Button buttonForMoveImagesByPath;
        private System.Windows.Forms.TextBox textBoxOfPathForMoveImage;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBoxOfResultForMoveImage;
        private System.Windows.Forms.Button buttonForSpareImages;
        private System.Windows.Forms.Button buttonForLostImages;
        private System.Windows.Forms.Button buttonForSQL;
        private System.Windows.Forms.Button buttonForCheck;
    }
}

