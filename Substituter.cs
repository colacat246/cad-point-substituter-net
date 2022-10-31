/* ==============================================================================
 * 功能描述：$substitute cad mp data$  
 * 创 建 者：jyj
 * 创建日期：$time$
 * CLR Version :$clrversion$
 * ==============================================================================*/

// 引用库文件
// 系统库
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections;
using System.Text;
using System.Threading.Tasks;
// 此库用于弹出文件选择对话框
using Wnd = System.Windows.Forms; 
// Cad提供的API
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace cad_point_substituter_net
{
    public class Substituter
    {
        [CommandMethod("SubMP")] // 在cad命令行中输入submp即可调用此函数
        public async void SubMP() // 测点替换函数
        {
            //System.Diagnostics.Stopwatch watch = new System.Diagnostics.Stopwatch(); // 启动代码运行时间监视

            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor; // 拿到cad命令行对象

            string dwgPath = SelectDwg(); // 选择Dwg文件，具体见下方的SelectDwg函数
            string[] txtPaths = SelectTxt(); // 选择测点替换文件，具体见SelectTxt函数
            string outputPath = SelectOutPutPath(); // 选择输出路径，具体见SelectOutPutPath函数，注意拿到的路径没有最后一个反斜杠

            if(dwgPath == null || txtPaths == null || outputPath == null) // 检查路径是否为空
            {
                ed.WriteMessage("文件或输出路径未选择"); // 是空则提示信息并返回，不执行之后的操作
                return;
            }

            // watch.Start();  //开始监视代码运行时间

            ArrayList log = new ArrayList(); // 初始化一个动态数组保存日志
            log.Add("OprationTimeStamp:" + DateTime.Now.ToString()); // 写入日志头
            log.Add("DwgFile:" + dwgPath); // 写入所使用的dwg文件
            foreach (string txtFile in txtPaths) // 写入所使用的测点替换文件
            {
                log.Add("SubstituteDataFiles:" + txtFile);

            }
            log.Add("OutputPath:" + outputPath); // 写入输出文件路径

            foreach (string txtFile in txtPaths) // 对每个文件替换测点
            {
                // 用字符串分割解析替换文件的文件名，以写入日志前缀
                string[] txtPathSplited = txtFile.Split('\\');
                int txtPathLen = txtPathSplited.Length - 1;

                string[] txtNameSplited = txtPathSplited[txtPathLen].Split('.');
                string txtName = txtNameSplited[0];

                (Dictionary<string, string[]> data, int angleTotal) = ReadInput(txtFile); // 读取测点替换文件，返回两个值
                                                                                          // 第一个值为字典，它的键为测点号，值为各风向角系数值的数组
                                                                                          // 第二个值为总风向角个数
                                                                                          // 读取过程见ReadInput函数

                int angleNum = 0; // 设定初始风向角
                while (true) // 对每个风向角循环操作
                {
                    string logPrefix = "data:" + txtName + "||angle:" + angleNum + "||"; // 日志前缀

                    // 声明图形数据库对象
                    Document doc = Application.DocumentManager.Add(dwgPath); // 加载文件，并拿到文档对象
                    DocumentLock docLock = doc.LockDocument(); // 锁定文档
                    Database db = doc.Database; // 从文档对象拿到数据库对象

                    using (Transaction trans = db.TransactionManager.StartTransaction()) // 开启数据库连接，得到连接对象trans 使用using关键字可在用完后自动关闭数据库连接，类似于python中with open的效果
                    {
                        BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead); // 通过trans对象的GetObject方法拿到块表，并向下转型，以调用BlockTable的方法
                        BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite); // 通过块表拿到模型空间的块表记录对象btr，这个btr对象可迭代，每次迭代返回模型空间中的元素id

                        foreach (ObjectId id in btr) // 遍历btr，对每个id进行操作
                        {
                            DBObject ent = trans.GetObject(id, OpenMode.ForWrite); // 通过元素id拿到模型空间中的元素对象ent

                            if (ent is DBText) // 判断ent是否为单行文字DBText类
                            { 
                                DBText dbText = (DBText)ent; // 如果是，则向下转型，以调用DBText的方法
                                
                                if (data.ContainsKey(dbText.TextString)) // 判断单行文字元素的值记录在TextString属性中，此处判断文字值是否在测点替换表中，如果在，则进行替换
                                {
                                    log.Add(logPrefix + "before:" + dbText.TextString + "||after:" + data[dbText.TextString][angleNum]); // 在日志中记录替换前后的值
                                    dbText.TextString = dbText.TextString.Replace(dbText.TextString, data[dbText.TextString][angleNum]); // 调用DBText.TextReplace.Replace方法进行替换
                                }
                                else // 不再测点替换表中则不替换
                                {
                                    log.Add(logPrefix + ">>>" + dbText.TextString + "<<< not substituted"); // 在日志中记录未替换的点
                                }
                            }
                        }
                        trans.Commit(); // 全部元素都识别并替换结束后提交transaction，保存替换完成后的数据
                    }

                    // 保存并关闭文件
                    db.SaveAs(outputPath + "\\" + txtName + "_" + angleNum.ToString() + ".dwg", DwgVersion.AC1024); // 可以选版本号：DwgVersion.Current -> 输出版本不变，DwgVersion.AC1024 -> Cad2010版
                    docLock.Dispose();  // 释放文档锁
                    doc.CloseAndDiscard(); // 关闭当前文件，不保存

                    angleNum += 1; // 处理完一个风向角，计数+1

                    if (angleNum == angleTotal) break; // 判断是否完成全部风向角，完成则停止循环
                }
            }

            // 输出日志
            string[] logTotal = (string[])log.ToArray(typeof(string));
            File.WriteAllLines(outputPath + "\\SubstitutionLog-" + DateTime.Now.ToString().Replace("/", "-").Replace(" ", "--").Replace(":", "-") + ".txt", logTotal);
            // watch.Stop();  //停止监视
            // TimeSpan timespan = watch.Elapsed;  //获取当前实例测量得出的总时间
            // ed.WriteMessage("测点替换完成，执行时间：{0}(毫秒)", timespan.TotalMilliseconds);  //总毫秒数
            ed.WriteMessage("测点替换完成");  
        }

        string SelectDwg() // 此函数用于选择Dwg文件路径，返回Dwg文件路径
        {
            Wnd.OpenFileDialog openDlg = new Wnd.OpenFileDialog(); // 调用系统的文件对话框API，弹出选择文件对话框
            openDlg.Title = "选择Cad文件"; // 指定对话框标题
            openDlg.Filter = "CAD文件(*.dwg)|*.dwg"; // 指定对话框能选择的文件
            Wnd.DialogResult openRes = openDlg.ShowDialog(); // 拿到选择的文件路径

            if (openRes == Wnd.DialogResult.OK) // 如果用户正常选择了路径，则返回路径
            {
                return openDlg.FileName;
            }
            else // 否则返回空
            {
                return null;
            }

        }

        string[] SelectTxt() // 弹出打开文件对话框，返回文件路径数组，原理同上
        {
            Wnd.OpenFileDialog openDlg = new Wnd.OpenFileDialog();
            openDlg.Multiselect = true; // 允许同时选择多个文件
            openDlg.Title = "选择测点替换文件";
            // openDlg.Filter = "文本文件(*.txt)|*.txt";
            Wnd.DialogResult openRes = openDlg.ShowDialog();
            if (openRes == Wnd.DialogResult.OK)
            {
                return openDlg.FileNames;
            }
            else
            {
                return null;
            }

        }

        string SelectOutPutPath() // 弹出打开文件对话框，返回输出文件夹，原理同上
        {
            Wnd.FolderBrowserDialog openDlg = new Wnd.FolderBrowserDialog();
            openDlg.Description = "请选择输出文件夹";
            Wnd.DialogResult openRes = openDlg.ShowDialog();
            if (openRes == Wnd.DialogResult.OK)
            {
                return openDlg.SelectedPath;
            }
            else
            {
                return null;
            }

        }


        (Dictionary<string, string[]>, int) ReadInput(string fileName) // 处理txt文件，返回字典，键为测点号，值为数组，以及总风向角
        {
            Dictionary<string, string[]> angleMap = new Dictionary<string, string[]>(); // 初始化一个字典，储存测点替换文件的数据
            ArrayList angleTotalList = new ArrayList(); // 新建一个数组，

            string[] contents = File.ReadAllLines(fileName); // 读取文件的所有行，返回内容数组

            foreach (string content in contents) // 遍历内容数组
            {
                string[] onePoint = Regex.Split(content.Trim(), "\\s+", RegexOptions.IgnoreCase); // 调用正则处理，用空格分割数据，返回数组
                string[] data = onePoint.Skip(1).ToArray(); // 拿到其中的风压系数数据
                angleTotalList.Add(data.Length); // 拿到每个测点的风压系数数据的长度，可以用来检测输入文件的一致性，这里没有检查
                string pointNum = onePoint[0]; // 拿到测点号数据
                angleMap.Add(pointNum, data); // 给字典添加 <测点号, 风压系数数据> 的键值对，
            }

            // TODO 检查每个测点的风压系数数据个数的一致性

            return (angleMap, (int)angleTotalList[0]); // 返回数据字典和总风向角个数

        }
    }
}
