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

namespace ImageManager
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonForFolderOfImages_Click(object sender, EventArgs e)
        {
            textBoxOfResultForCreateExcel.Clear();
            string path = GetImagPath();
            if (path == null) return;
            CreateExcel(path);
            textBoxOfResultForCreateExcel.AppendText(path + " 操作完成！\n");
        }

        private void buttonForPathOfImages_Click(object sender, EventArgs e)
        {
            textBoxOfResultForCreateExcel.Clear();
            string path = textBoxOfPathForCreateExcel.Text;
            if (path == null) return;
            CreateExcel(path);
            textBoxOfResultForCreateExcel.AppendText(path + " 操作完成！\n");
        }

        private void buttonForFolderOfExcels_Click(object sender, EventArgs e)
        {
            textBoxOfResultForCreateExcel.Clear();
            string[] excelPaths = GetExcelsPath();
            if (excelPaths == null) return;
            for(int i = 0; i < excelPaths.Length; i++)
            {
                OpenExcel(excelPaths[i]);
                textBoxOfResultForCreateExcel.AppendText(excelPaths[i] + " 操作完成！" + "\n");
            }
        }

        private void buttonForMoveImagesByFolder_Click(object sender, EventArgs e)
        {
            textBoxOfResultForMoveImage.Clear();
            string path = GetImagPath();
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择要保存的文件夹";
            dialog.ShowDialog();
            string storagePath = dialog.SelectedPath;
            if (path == null) return;
            string[] imagePaths = Directory.GetFiles(path, "*.jpg", SearchOption.AllDirectories);
            string[] fields = new string[5];
            for (int i = 0; i < imagePaths.Length; i++)
            {
                string imagName = Path.GetFileNameWithoutExtension(imagePaths[i]);
                //if (imagName.Contains("(") || imagName.Contains(")"))
                //{
                //    fields = imagName.Split(new string[] { "(", ")", "（", "）" }, 5, System.StringSplitOptions.None);
                //    if (fields[2] == null)
                //    {
                //        textBoxOfResultForMoveImage.AppendText("未移动" + imagName + "\n");
                //        continue;
                //    }
                //}
                //if (imagName.Contains("（") || imagName.Contains("）"))
                //{
                //    fields = imagName.Split(new string[] { "(", ")", "（", "）" }, 5, System.StringSplitOptions.None);
                //    if (string.IsNullOrEmpty(fields[2]))
                //    {
                //        textBoxOfResultForMoveImage.AppendText("未移动" + imagName + "\n");
                //        continue;
                //    }
                //}
                //if (imagName.EndsWith("0003") || imagName.EndsWith("0004"))
                //{
                //    textBoxOfResultForMoveImage.AppendText("出现未处理的" + imagName + "\n");
                //    continue;
                //}
                if (imagName.Contains("0001") && imagName.EndsWith("0001"))
                {
                   
                    //探查是否存在0002图片，如果存在则将该人写入excel，否则不写入
                    if( !File.Exists(imagePaths[i].Remove(imagePaths[i].Length - 8, 8) + "0002.jpg"))
                    {
                        continue;
                    }
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
                    if (imagName.Contains("(") || imagName.Contains(")"))
                    {
                        fields = imagName.Split(new string[] { "(", ")", "（", "）" }, 5, System.StringSplitOptions.None);
                    }
                    if (imagName.Contains("（") || imagName.Contains("）"))
                    {
                        fields = imagName.Split(new string[] { "(", ")", "（", "）" }, 5, System.StringSplitOptions.None);
                    }
                    if (fields == null)
                    {
                        fields = new string[4] { " ", " ", " ", " " };
                    }
                    if (fields[2] == "0001" | fields[2] == null)
                    {
                            
                        textBoxOfResultForMoveImage .AppendText("未移动" + imagName + "\n");
                        continue;
                    }
                }
                string code = GetCodes(imagePaths[i]);
                MoveImage(imagePaths[i], storagePath, code);
            }
            textBoxOfResultForMoveImage.AppendText(path + "图片转移完毕\n");

            }

        private void buttonForMoveImagesByPath_Click(object sender, EventArgs e)
        {
            textBoxOfResultForMoveImage.Clear();
            string path = textBoxOfPathForMoveImage.Text;
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择要保存的文件夹";
            dialog.ShowDialog();
            string storagePath = dialog.SelectedPath;
            if (path == null) return;
            string[] imagePaths = Directory.GetFiles(path, "*.jpg", SearchOption.AllDirectories);
            string[] fields = new string[5];
            for (int i = 0; i < imagePaths.Length; i++)
            {
                string imagName = Path.GetFileNameWithoutExtension(imagePaths[i]);
                if (imagName.Contains("(") || imagName.Contains(")"))
                {
                    fields = imagName.Split(new string[] { "(", ")", "（", "）" }, 5, System.StringSplitOptions.None);
                    if (fields[2] == null)
                    {
                        textBoxOfResultForMoveImage.AppendText("未移动" + imagName + "\n");
                        continue;
                    }
                }
                if (imagName.Contains("（") || imagName.Contains("）"))
                {
                    fields = imagName.Split(new string[] { "(", ")", "（", "）" }, 5, System.StringSplitOptions.None);
                    if (string.IsNullOrEmpty(fields[2]))
                    {
                        textBoxOfResultForMoveImage.AppendText("未移动" + imagName + "\n");
                        continue;
                    }
                }
                if (imagName.EndsWith("0003") || imagName.EndsWith("0004"))
                {
                    textBoxOfResultForMoveImage.AppendText("出现未处理的" + imagName + "\n");
                    continue;
                }
                if (imagName.Contains("0001") && imagName.EndsWith("0001"))
                {
                   
                    //探查是否存在0002图片，如果存在则将该人写入excel，否则不写入
                    if( !File.Exists(imagePaths[i].Remove(imagePaths[i].Length - 8, 8) + "0002.jpg"))
                    {
                        continue;
                    }
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
                    if (imagName.Contains("(") || imagName.Contains(")"))
                    {
                        fields = imagName.Split(new string[] { "(", ")", "（", "）" }, 5, System.StringSplitOptions.None);
                    }
                    if (imagName.Contains("（") || imagName.Contains("）"))
                    {
                        fields = imagName.Split(new string[] { "(", ")", "（", "）" }, 5, System.StringSplitOptions.None);
                    }
                    if (fields == null)
                    {
                        fields = new string[4] { " ", " ", " ", " " };
                    }
                    if (fields[2] == "0001" | fields[2] == null)
                    {
                            
                        textBoxOfResultForMoveImage .AppendText("未移动" + imagName + "\n");
                        continue;
                    }
                }
                string code = GetCodes(imagePaths[i]);
                MoveImage(imagePaths[i], storagePath, code);
            }
            textBoxOfResultForMoveImage.AppendText(path + "图片转移完毕\n");

        }

        private void buttonForSpareImages_Click(object sender, EventArgs e)
        {
            textBoxOfResultForCreateExcel.Clear();
            string[] imagePaths = GetFilePath();
            if (imagePaths == null) return;
            for(int i = 0; i <= imagePaths.Length - 1; i++)
            {
                if(imagePaths[i].EndsWith("0001.jpg") || imagePaths[i].EndsWith("0002.jpg"))
                {
                    continue;
                }
                else
                {
                    textBoxOfResultForCreateExcel.AppendText(imagePaths[i] + "\n");
                }
            }
            textBoxOfResultForCreateExcel.AppendText("检查完毕\n");
        }

        private void buttonForLostImages_Click(object sender, EventArgs e)
        {
            textBoxOfResultForCreateExcel.Clear();
            string[] imagePaths = GetFilePath();
            if (imagePaths == null) return;
            for (int i = 0; i <= imagePaths.Length - 1; i++)
            {
                string imageName = Path.GetFileNameWithoutExtension(imagePaths[i]);
                if (imageName.Contains("0001") && imageName.EndsWith("0001"))
                {
                    if (!File.Exists(imagePaths[i].Remove(imagePaths[i].Length - 8, 8) + "0002.jpg"))
                    {
                        textBoxOfResultForCreateExcel.AppendText(imagePaths[i] + "\n");
                    }
                }
            }
            textBoxOfResultForCreateExcel.AppendText("检查完毕\n");
        }

        private void buttonForSQL_Click(object sender, EventArgs e)
        {
            textBoxOfResultForCreateExcel.Clear();

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.DefaultExt = ".xlsx";
            dialog.Title = "请选择目标文件";
            dialog.ShowDialog();
            string excelPath = dialog.FileName;
            if (excelPath == null) return;
            textBoxOfPathForCreateExcel.AppendText("已选中" + excelPath + "\n");
            bool issucceed = LoadExcelData(excelPath);
            if (issucceed) textBoxOfResultForCreateExcel.AppendText("\n数据导入完毕！\n");
            else textBoxOfResultForCreateExcel.AppendText("数据导入失败！\n");
        }

        private void buttonForCheck_Click(object sender, EventArgs e)
        {
            textBoxOfPathForCreateExcel.Clear();
            textBoxOfResultForCreateExcel.Clear();
            string path = GetImagPath();
            if (path == null) return;
            textBoxOfPathForCreateExcel.AppendText("已选中：" + path);
            CheckImageInfo(path);
        }
    }
}
