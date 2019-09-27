using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using NPOI.XWPF.UserModel;
using System.IO;

namespace WordUtility_for_NPOI
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            List<string> ComboData = new List<string>() { "仇丽颖", "韦松", "莫宝津", "杨海霞", "宋睿", "胡月媛", "赵康康", "范志远", "李永伟", "邱太福", "王成辉", "江双禧", "罗振兴", "刘庚铖", "余成鑫", "金海康", "赖振东", "陈兵", "何培佳", "陈河棋", "张建", "揭腾辉", "张再瀛", "赵景旭", "袁东", "庄泽钦" };//ComboBox下拉选项            
            ComboBox.ItemsSource = ComboData;
        }


        public class qslyclass
        {
            public int Para { get; set; }

            public int Runs { get; set; }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog folderDlg = new OpenFileDialog();
            folderDlg.Filter = "模板文件|*.docx";
            folderDlg.ShowDialog();
            string filename = folderDlg.FileName;
            //FolderBrowserDialog.Description = "请选择数据库路径";
            //FolderBrowserDialog.ShowDialog();
            TextBox.Text = "" + filename;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            OpenFileDialog folderDlg = new OpenFileDialog();
            folderDlg.Filter = "ACESS文件|*.mdb";
            folderDlg.ShowDialog();
            string filename = folderDlg.FileName;
            //FolderBrowserDialog.Description = "请选择数据库路径";
            //FolderBrowserDialog.ShowDialog();
            TextBox2.Text = "" + filename;
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            OpenFileDialog folderDlg = new OpenFileDialog();
            folderDlg.Filter = "ACESS文件|*.mdb";
            folderDlg.ShowDialog();
            string filename = folderDlg.FileName;
            //FolderBrowserDialog.Description = "请选择数据库路径";
            //FolderBrowserDialog.ShowDialog();
            TextBox3.Text = "" + filename;
        }

       

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            #region 变量
            string surveyNO = "";//调查表编号
            string streetName = "";//街道名称
            string parcelNO = "";//宗地号
            string parcelName = "";//宗地名称
            string tdzl = "";//土地坐落
            string qlr = "";//权利人
            string qsly = "";//权属来源
            string qslyNO = "";//权属来源文号
            string jsUser = "";//调查记事实际使用人
            //string jsUserOption = "";//调查记事实际使用人选项
            string jsUse = "";//调查记事实际用途/产业
            string jsChange = "";//调查记事建筑是否发生变化
            string jsOther = "";//调查记事其他
            string jsPerson = "";//调查员
            string jsTime = "";//调查时间

            #endregion

            //弹出保存文件对话框，保存生成的Word
            FolderBrowserDialog FolderBrowserDialog = new FolderBrowserDialog();
            FolderBrowserDialog.Description = "请选择文件路径";
            FolderBrowserDialog.ShowDialog();


            string dotpath = TextBox.Text;
            string path = TextBox2.Text;
            string path2 = TextBox3.Text;
            string street = StreetText.Text;
            //string excname = ExcelText.Text;

            string parsql = "SELECT exc.门牌号, exc.权利人, exc.[实际用途/产业], exc.[记事-外业核实记录表], exc.建筑是否发生变化, exc.宗地名称, par.SURVEY_NO, par.宗地号_1 " +
                "FROM LAND_CADASTRAL_SURVEY_PARCEL par LEFT JOIN Sheet1 exc on par.宗地号_1=exc.宗地号";
            var parTable = OleDbHelper.QueryTable(parsql, path);

            string tbsql = "SELECT FILE_ID, PARCEL_NO, VALID_FLAG, FILE_TYPE, APPROVE_NO FROM LAND_QSLYWJ_TB";
            var tbTable = OleDbHelper.QueryTable(tbsql, path2);

            string tbownersql = "SELECT FILE_ID, NAME FROM LAND_QSLYWJ_OWNER";
            var toTable = OleDbHelper.QueryTable(tbownersql, path2);

            int res = 1;

            for (int i = 0; i < parTable.Rows.Count; i++)
            {
                List<qslyclass> qslylist = new List<qslyclass>();

                surveyNO = parTable.Rows[i]["SURVEY_NO"].ToString();
                parcelNO = parTable.Rows[i]["宗地号_1"].ToString();
                streetName = street;
                parcelName = parTable.Rows[i]["宗地名称"].ToString();
                tdzl = "罗湖区" + street + parTable.Rows[i]["门牌号"].ToString();

                var tbrow = tbTable.Select("PARCEL_NO = '" + parTable.Rows[i]["宗地号_1"].ToString() + "' AND VALID_FLAG = '1'");
                if (tbrow.Length != 0)
                {
                    for (int j = 0; j < tbrow.Length; j++)
                    {
                        //权利人
                        var tomanerow = toTable.Select("FILE_ID = '" + tbrow[j]["FILE_ID"].ToString() + "'");
                        if (tomanerow.Length != 0)
                        {
                            string toname = tomanerow[0]["NAME"].ToString();
                            if (qlr == "")
                            {
                                qlr = toname;
                            }
                            else
                            {
                                if (qlr.Contains(toname))
                                {

                                }
                                else
                                {
                                    qlr = qlr + "，" + toname;
                                }
                            }
                        }
                        //权属来源
                        string tbtype = "";
                        int qsly1 = -1;
                        int qsly2 = -1;
                        qslyclass qslyclass = new qslyclass();

                        switch (tbrow[j]["FILE_TYPE"].ToString())
                        {
                            case "不动产权利证书":
                                tbtype = "20";
                                break;
                            case "土地使用合同书及付清地价款证明":
                                tbtype = "21";
                                break;
                            case "政府批准用地文件及用地红线图（划拨决定书）":
                                tbtype = "22";
                                break;
                            case "政府批地会批复（纪要）及用地方案图":
                                tbtype = "23";
                                break;
                            case "非农建设用地等原农村用地批准文件":
                                tbtype = "24";
                                break;
                            case "历史遗留问题处理决定书":
                                tbtype = "25";
                                break;
                            case "政府法院仲裁机构用地处理结果法律文书":
                                tbtype = "26";
                                break;
                            case "国有未出让土地划线移交或委托管理文件":
                                tbtype = "27";
                                break;
                            case "以其他合法形式取得不动产权利的证明文件":
                                tbtype = "28";
                                break;
                            case "征转收地协议":
                                tbtype = "29";
                                break;
                            case "土地权属界线协议书":
                                tbtype = "30";
                                break;
                            case "土地权属争议(异议)原由书":
                                tbtype = "31";
                                break;
                            case "已有地籍调查成果资料":
                                tbtype = "32";
                                break;
                            case "地籍调查资料查询结果证明":
                                tbtype = "33";
                                break;
                            case "建设用地规划许可证":
                                tbtype = "34";
                                break;
                            case "建设工程规划许可证":
                                tbtype = "35";
                                break;
                            case "建设工程规划验收合格证":
                                tbtype = "36";
                                break;
                            case "用地用房信息申报表":
                                tbtype = "37";
                                break;
                            case "土地权属演变情况说明":
                                tbtype = "38";
                                break;
                            case "地籍调查任务书":
                                tbtype = "39";
                                break;
                        }
                        switch (tbtype)
                        {
                            case "20":
                                qsly1 = 0;
                                qsly2 = 0;
                                break;
                            case "21":
                                qsly1 = 0;
                                qsly2 = 2;
                                break;
                            case "22":
                                qsly1 = 0;
                                qsly2 = 4;
                                break;
                            case "23":
                                qsly1 = 0;
                                qsly2 = 6;
                                break;
                            case "24":
                                qsly1 = 0;
                                qsly2 = 8;
                                break;
                            case "25":
                                qsly1 = 0;
                                qsly2 = 10;
                                break;
                            case "26":
                                qsly1 = 0;
                                qsly2 = 12;
                                break;
                            case "27":
                                qsly1 = 0;
                                qsly2 = 14;
                                break;
                            case "28":
                                qsly1 = 0;
                                qsly2 = 16;
                                break;
                            case "29":
                                qsly1 = 0;
                                qsly2 = 18;
                                break;
                            case "30":
                                qsly1 = 1;
                                qsly2 = 0;
                                break;
                            case "31":
                                qsly1 = 1;
                                qsly2 = 2;
                                break;
                            case "32":
                                qsly1 = 1;
                                qsly2 = 4;
                                break;
                            case "33":
                                qsly1 = 1;
                                qsly2 = 6;
                                break;
                            case "34":
                                qsly1 = 1;
                                qsly2 = 8;
                                break;
                            case "35":
                                qsly1 = 1;
                                qsly2 = 10;
                                break;
                            case "36":
                                qsly1 = 1;
                                qsly2 = 12;
                                break;
                            case "37":
                                qsly1 = 1;
                                qsly2 = 14;
                                break;
                            case "38":
                                qsly1 = 1;
                                qsly2 = 16;
                                break;
                            case "39":
                                qsly1 = 1;
                                qsly2 = 18;
                                break;
                        }
                        qslyclass.Para = qsly1;
                        qslyclass.Runs = qsly2;
                        qslylist.Add(qslyclass);
                        



                        //权属来源文号
                        string tbno = tbrow[j]["APPROVE_NO"].ToString();
                        if (qslyNO == "")
                        {
                            qslyNO = tbno;
                        }
                        else
                        {
                            if (qslyNO.Contains(tbno))
                            {

                            }
                            else
                            {
                                qslyNO = qslyNO + "，" + tbno;
                            }
                        }
                    }


                }

                jsUser = parTable.Rows[i]["权利人"].ToString();
                jsUse = parTable.Rows[i]["实际用途/产业"].ToString();
                jsChange = parTable.Rows[i]["建筑是否发生变化"].ToString();
                if (String.IsNullOrEmpty(jsChange))
                {
                    jsChange = "建筑物无变化";
                }
                jsOther = parTable.Rows[i]["记事-外业核实记录表"].ToString();
                jsPerson = ComboBox.Text;
                var time = Convert.ToDateTime(DatePicker.SelectedDate);
                Random rd = new Random();
                int day = rd.Next(1, 30);
                jsTime = time.Year.ToString() + "年" + time.Month.ToString() + "月" + day.ToString() + "日";



                #region word文档
                try
                {
                    XWPFDocument myDocx = null;
                    FileStream fs = null;
                    fs = new FileStream(TextBox.Text, FileMode.Open, FileAccess.Read);      
                    myDocx = new XWPFDocument(fs);//打开docx




                    var para1 = myDocx.Tables[0].Rows[0].GetTableCells()[1].Paragraphs[0];
                    string oldtext1 = para1.ParagraphText;
                    string temptext1 = "";
                    if (oldtext1.Contains("{$surveyNO}"))
                        temptext1 = oldtext1.Replace("{$surveyNO}", surveyNO);
                    para1.ReplaceText(oldtext1, temptext1);
                   
                    var para2 = myDocx.Tables[0].Rows[0].GetTableCells()[3].Paragraphs[0];
                    string oldtext2 = para2.ParagraphText;
                    string temptext2 = "";
                    if (oldtext2.Contains("{$streetName}"))
                        temptext2 = oldtext2.Replace("{$streetName}",streetName);
                    para2.ReplaceText(oldtext2, temptext2);

                    var para3 = myDocx.Tables[0].Rows[1].GetTableCells()[1].Paragraphs[0];
                    string oldtext3 = para3.ParagraphText;
                    string temptext3 = "";
                    if (oldtext3.Contains("{$parcelNO}"))
                        temptext3 = oldtext3.Replace("{$parcelNO}", parcelNO);
                    para3.ReplaceText(oldtext3, temptext3);

                    var para4 = myDocx.Tables[0].Rows[1].GetTableCells()[3].Paragraphs[0];
                    string oldtext4 = para4.ParagraphText;
                    string temptext4 = "";
                    if (oldtext4.Contains("{$parcelName}"))
                        temptext4 = oldtext4.Replace("{$parcelName}", parcelName);
                    para4.ReplaceText(oldtext4, temptext4);

                    var para5 = myDocx.Tables[0].Rows[2].GetTableCells()[1].Paragraphs[0];
                    string oldtext5 = para5.ParagraphText;
                    string temptext5 = "";
                    if (oldtext5.Contains("{$tdzl}"))
                        temptext5 = oldtext5.Replace("{$tdzl}", tdzl);
                    para5.ReplaceText(oldtext5, temptext5);

                    var para6 = myDocx.Tables[0].Rows[3].GetTableCells()[1].Paragraphs[0];
                    string oldtext6 = para6.ParagraphText;
                    string temptext6 = "";
                    if (oldtext6.Contains("{$qlr}"))
                        temptext6 = oldtext6.Replace("{$qlr}", qlr);
                    para6.ReplaceText(oldtext6, temptext6);


                    //var para7 = myDocx.Tables[0].Rows[4].GetTableCells()[1].Paragraphs[0];
                    //string oldtext7 = para7.ParagraphText;
                    //string temptext7 = "";
                    //if (oldtext7.Contains("{$qsly}"))
                    //temptext7 = oldtext7.Replace("{$qsly}", qsly);
                    //para7.ReplaceText(oldtext7, temptext7);
                    if (qslylist.Count != 0)
                    {
                        for (int j = 0; j < qslylist.Count; j++)
                        {
                            if (qslylist[j].Para != -1)
                            {
                                var para7 = myDocx.Tables[0].Rows[4].GetTableCells()[1].Paragraphs[qslylist[j].Para];
                                var a = para7.Runs[qslylist[j].Runs];
                                a.FontFamily = "Wingdings 2";
                                a.SetText("R");
                            }
                        }
                    }


                    var para8 = myDocx.Tables[0].Rows[5].GetTableCells()[1].Paragraphs[0];
                    string oldtext8 = para8.ParagraphText;
                    string temptext8 = "";
                    if (oldtext8.Contains("{$qslyNO}"))
                        temptext8 = oldtext8.Replace("{$qslyNO}", qslyNO);
                    para8.ReplaceText(oldtext8, temptext8);

                    var para9 = myDocx.Tables[0].Rows[6].GetTableCells()[1].Paragraphs[0];
                    string oldtext9= para9.ParagraphText;
                    string temptext9 = "";
                    if (oldtext9.Contains("{$jsUser}"))
                        temptext9 = oldtext9.Replace("{$jsUser}", jsUser);
                    para9.ReplaceText(oldtext9, temptext9);

                    //var para10 = myDocx.Tables[0].Rows[7].GetTableCells()[1].Paragraphs[1];
                    //string oldtext10 = para10.ParagraphText;
                    //string temptext10 = "";
                    //if (oldtext10.Contains("{$jsUse}"))
                    //    temptext10= oldtext10.Replace("{$jsUse}", jsUse);
                    //para10.ReplaceText(oldtext10, temptext10);
                    if (jsUse.Length != 0)
                    {
                        if (jsUse.Contains("工业"))
                        {
                            var para10 = myDocx.Tables[0].Rows[7].GetTableCells()[1].Paragraphs[0];
                            var b = para10.Runs[1];
                            b.FontFamily = "Wingdings 2";
                            b.SetText("R");
                        }
                        if (jsUse.Contains("商服"))
                        {
                            var para10 = myDocx.Tables[0].Rows[7].GetTableCells()[1].Paragraphs[0];
                            var b = para10.Runs[5];
                            b.FontFamily = "Wingdings 2";
                            b.SetText("R");
                        }
                        if (jsUse.Contains("居住"))
                        {
                            var para10 = myDocx.Tables[0].Rows[7].GetTableCells()[1].Paragraphs[0];
                            var b = para10.Runs[3];
                            b.FontFamily = "Wingdings 2";
                            b.SetText("R");
                            var c = para10.Runs[4];
                            c.SetText(jsUse);
                        }
                        if (jsUse.Contains("其他"))
                        {
                            var para10 = myDocx.Tables[0].Rows[7].GetTableCells()[1].Paragraphs[0];
                            var b = para10.Runs[7];
                            b.FontFamily = "Wingdings 2";
                            b.SetText("R");
                            var c = para10.Runs[8];
                            c.SetText(jsUse);
                        }
                    }
                    


                    var para11 = myDocx.Tables[0].Rows[8].GetTableCells()[1].Paragraphs[0];
                    string oldtext11 = para11.ParagraphText;
                    string temptext11 = "";
                    if (oldtext11.Contains("{$jsChange}"))
                        temptext11 = oldtext11.Replace("{$jsChange}", jsChange);
                    para11.ReplaceText(oldtext11, temptext11);

                    var para12 = myDocx.Tables[0].Rows[9].GetTableCells()[1].Paragraphs[0];
                    string oldtext12 = para12.ParagraphText;
                    string temptext12 = "";
                    if (oldtext12.Contains("{$jsOther}"))
                        temptext12 = oldtext12.Replace("{$jsOther}", jsOther);
                    para12.ReplaceText(oldtext12, temptext12);

                    var para13 = myDocx.Tables[0].Rows[11].GetTableCells()[1].Paragraphs[0];
                    string oldtext13 = para13.ParagraphText;
                    string temptext13 = "";
                    if (oldtext13.Contains("{$jsPerson}"))
                        temptext13 = oldtext13.Replace("{$jsPerson}", jsPerson);
                    para13.ReplaceText(oldtext13, temptext13);

                    var para14 = myDocx.Tables[0].Rows[11].GetTableCells()[3].Paragraphs[0];
                    string oldtext14 = para14.ParagraphText;
                    string temptext14 = "";
                    if (oldtext14.Contains("{$jsTime}"))
                        temptext14 = oldtext14.Replace("{$jsTime}", jsTime);
                    para14.ReplaceText(oldtext14, temptext14);

                    string filename = FolderBrowserDialog.SelectedPath + "\\权利人现场核实记录表（" + parcelNO + "）.docx";

                    FileStream output = new FileStream(filename, FileMode.Create);
                    myDocx.Write(output);

                    fs.Close();
                    fs.Dispose();
                    output.Close();
                    output.Dispose();

                    //清空变量
                    surveyNO = "";//调查表编号
                    streetName = "";//街道名称
                    parcelNO = "";//宗地号
                    parcelName = "";//宗地名称
                    tdzl = "";//土地坐落
                    qlr = "";//权利人
                    qsly = "";//权属来源
                    qslyNO = "";//权属来源文号
                    jsUser = "";//调查记事实际使用人
                    //string jsUserOption = "";//调查记事实际使用人选项
                    jsUse = "";//调查记事实际用途/产业
                    jsChange = "";//调查记事建筑是否发生变化
                    jsOther = "";//调查记事其他
                    jsPerson = "";//调查员
                    jsTime = "";//调查时间
                    qslylist.Clear();
                    res++;

                }
                catch(Exception ex)
                {
                    System.Windows.MessageBox.Show(parcelNO);
                    //清空变量
                    surveyNO = "";//调查表编号
                    streetName = "";//街道名称
                    parcelNO = "";//宗地号
                    parcelName = "";//宗地名称
                    tdzl = "";//土地坐落
                    qlr = "";//权利人
                    qsly = "";//权属来源
                    qslyNO = "";//权属来源文号
                    jsUser = "";//调查记事实际使用人
                    //string jsUserOption = "";//调查记事实际使用人选项
                    jsUse = "";//调查记事实际用途/产业
                    jsChange = "";//调查记事建筑是否发生变化
                    jsOther = "";//调查记事其他
                    jsPerson = "";//调查员
                    jsTime = "";//调查时间
                    qslylist.Clear();
                    System.Windows.MessageBox.Show(ex.ToString());
                    break;
                }
                #endregion
            }
            System.Windows.Forms.MessageBox.Show("成功生成" + res + "项");
        }



       
        
    }
}
