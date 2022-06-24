using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Wnd = System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections;

namespace cad_point_substituter_net
{
    public class Substituter
    {
        [CommandMethod("SubMP")]
        public async void SubMP()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            // 选择Dwg文件
            string dwgPath = SelectDwg();
            // 选择测点替换文件
            string[] txtPaths = SelectTxt();
            // 选择输出路径            
            string outputPath = SelectOutPutPath(); // 保存的路径没有最后一个反斜杠

            // 日志
            ArrayList log = new ArrayList();
            log.Add("OprationTimeStamp:" + DateTime.Now.ToString());
            log.Add("DwgFile:" + dwgPath);
            foreach (string txtFile in txtPaths)
            {
                log.Add("SubstituteDataFiles:" + txtFile);

            }
            log.Add("OutputPath:" + outputPath);

            // 对每个文件替换测点
            foreach (string txtFile in txtPaths)
            {
                // 分割txt文件名
                string[] txtPathSplited = txtFile.Split('\\');
                int txtPathLen = txtPathSplited.Length - 1;

                string[] txtNameSplited = txtPathSplited[txtPathLen].Split('.');
                string txtName = txtNameSplited[0];

                (Dictionary<string, string[]> data, int angleTotal) = ReadInput(txtFile);
                int angleNum = 0;
                while (true)
                {
                    // 日志前缀
                    string logPrefix = "data:" + txtName + "||angle:" + angleNum + "||";

                    // 声明图形数据库对象
                    Document doc = Application.DocumentManager.Add(dwgPath);
                    DocumentLock docLock = doc.LockDocument();
                    Database db = doc.Database;

                    using (Transaction trans = db.TransactionManager.StartTransaction())
                    {// 用完后自动关闭，类似于with open
                        // 打开块表
                        BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        // 打开块表记录
                        BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        // 加直线到块表记录
                        foreach (ObjectId id in btr)
                        {
                            // 获取数据库中的对象
                            DBObject ent = trans.GetObject(id, OpenMode.ForWrite);
                            // 判断对象是否为单行文字
                            if (ent is DBText)
                            {
                                DBText dbText = (DBText)ent;
                                // 替换文本
                                if (data.ContainsKey(dbText.TextString))
                                {
                                    log.Add(logPrefix + "before:" + dbText.TextString + "||after:" + data[dbText.TextString][angleNum]);
                                    dbText.TextString = dbText.TextString.Replace(dbText.TextString, data[dbText.TextString][angleNum]);
                                }
                                else
                                {
                                    log.Add(logPrefix + ">>>" + dbText.TextString + "<<< not substituted");
                                }
                            }
                        }
                        trans.Commit();


                    }

                    // 保存并关闭文件
                    db.SaveAs(outputPath + "\\" + txtName + "_" + angleNum.ToString() + ".dwg", DwgVersion.Current);
                    docLock.Dispose();
                    doc.CloseAndDiscard();

                    angleNum += 1;
                    // 判断是否完成全部风向角
                    if (angleNum == angleTotal) break;
                }
            }

            // 输出日志
            string[] logTotal = (string[])log.ToArray(typeof(string));
            File.WriteAllLines(outputPath + "\\SubstitutionLog-" + DateTime.Now.ToString().Replace("/", "-").Replace(" ", "--").Replace(":", "-") + ".txt", logTotal);
        }

        // 弹出打开文件对话框，返回Dwg文件路径
        string SelectDwg()
        {
            Wnd.OpenFileDialog openDlg = new Wnd.OpenFileDialog();
            openDlg.Title = "选择Cad文件";
            openDlg.Filter = "CAD文件(*.dwg)|*.dwg";
            Wnd.DialogResult openRes = openDlg.ShowDialog();
            if (openRes == Wnd.DialogResult.OK)
            {
                return openDlg.FileName;
            }
            else
            {
                return null;
            }

        }

        // 弹出打开文件对话框，返回文件路径数组
        string[] SelectTxt()
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

        // 弹出打开文件对话框，返回输出文件夹
        string SelectOutPutPath()
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


        // 处理txt文件，返回字典，键为测点号，值为数组，以及总风向角
        (Dictionary<string, string[]>, int) ReadInput(string fileName)
        {
            Dictionary<string, string[]> angleMap = new Dictionary<string, string[]>();
            ArrayList angleTotalList = new ArrayList();
            string[] contents = File.ReadAllLines(fileName);
            foreach (string content in contents)
            {
                string[] onePoint = Regex.Split(content.Trim(), "\\s+", RegexOptions.IgnoreCase);
                string[] data = onePoint.Skip(1).ToArray();
                angleTotalList.Add(data.Length);
                string pointNum = onePoint[0];
                angleMap.Add(pointNum, data);
            }
            // 差一个检查个测点风向角个数的操作
            return (angleMap, (int)angleTotalList[0]);

        }
    }
}
