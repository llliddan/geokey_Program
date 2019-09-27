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
using System.Data;

namespace QLRUpdate
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog folderDlg = new OpenFileDialog();
            folderDlg.Filter = "ACCESS|*.mdb";
            folderDlg.ShowDialog();
            string filename = folderDlg.FileName;
            //FolderBrowserDialog.Description = "请选择数据库路径";
            //FolderBrowserDialog.ShowDialog();
            TextBox1.Text = "" + filename;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog folderDlg = new OpenFileDialog();
            folderDlg.Filter = "ACCESS|*.mdb";
            folderDlg.ShowDialog();
            string filename = folderDlg.FileName;
            //FolderBrowserDialog.Description = "请选择数据库路径";
            //FolderBrowserDialog.ShowDialog();
            TextBox3.Text = "" + filename;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            try
            {
                string path = TextBox1.Text;
                string path2 = TextBox3.Text;
                string path3 = TextBox3_Copy.Text;
                string jdname = TextBox2.Text;

                #region 修改记事
                //删除记事
                string PARCEL_dejs = "UPDATE LAND_CADASTRAL_SURVEY_PARCEL " +
                    "SET LAND_CADASTRAL_SURVEY_PARCEL.REMARK = '', LAND_CADASTRAL_SURVEY_PARCEL.OWNER_INFO = '', LAND_CADASTRAL_SURVEY_PARCEL.OWNERSHIP_SURVEY = '', LAND_CADASTRAL_SURVEY_PARCEL.SURVEY_RECORD = '';";
                OleDbHelper.RunCommand(PARCEL_dejs, path);

                //
                string Parcel_NFO = "UPDATE LAND_CADASTRAL_SURVEY_PARCEL SET LAND_OTHERRIGHT_INFO ='无'";
                OleDbHelper.RunCommand(Parcel_NFO, path);

                //修改记事
                //string PARCEL_js = "UPDATE LAND_CADASTRAL_SURVEY_PARCEL AS a, " + jsname + " AS b SET a.REMARK = [b].[记事], a.OWNERSHIP_SURVEY = [b].[权属调查记事], a.SURVEY_RECORD = [b].[地籍测量记事] WHERE (a.PARCEL_NO) =[b].[宗地号]; ";
                string PARCEL_js = "UPDATE LAND_CADASTRAL_SURVEY_PARCEL as a, Sheet1 as b SET a.REMARK = b.记事, a.OWNERSHIP_SURVEY = b.权属调查记事, a.SURVEY_RECORD = b.地籍测量记事, a.LU_LOCATION = '"+ jdname + "' + b.门牌号, a.PARCEL_NAME = b.宗地名称 " +
                "WHERE a.宗地号_1 = b.宗地号";
                //string PARCEL_js = "UPDATE LAND_CADASTRAL_SURVEY_PARCEL as a, " + jsname + " as b SET a.REMARK = b.记事 " +
                //"WHERE a.PARCEL_NO = b.宗地号; ";
                int r2 = OleDbHelper.RunCommand(PARCEL_js, path);
                if (r2 == 0)
                {
                    System.Windows.Forms.MessageBox.Show("修改记事失败");
                }
                #endregion

                string OWNER = "SELECT * FROM LAND_CADASTRAL_SURVEY_OWNER";
                var OWNERTable = OleDbHelper.QueryTable(OWNER, path);

                string PARCEL = "SELECT exc.权利人, par.OWNER_INFO, par.PARCEL_CODE, par.OBJECTID, par.宗地号_1 " +
                    "FROM LAND_CADASTRAL_SURVEY_PARCEL as par LEFT JOIN Sheet1 as exc on par.宗地号_1 = exc.宗地号";
                var PARCELTable = OleDbHelper.QueryTable(PARCEL, path);

                string PARCEL2 = "SELECT * FROM LAND_CADASTRAL_SURVEY_PARCEL";
                var PARCELTable2 = OleDbHelper.QueryTable(PARCEL2, path2);

                string js = "SELECT * FROM Sheet1";
                var JSTable = OleDbHelper.QueryTable(js, path);

                
                






                #region 修改权利来源
                //for (int i = 0; i < OWNERTable.Rows.Count; i++)
                //{
                //    string CARD_NO = OWNERTable.Rows[i]["CARD_NO"].ToString();
                //    if (CARD_NO != "" && CARD_NO != null)
                //    {
                //        if (OWNER_INFO == "")
                //        {
                //            OWNER_INFO = "12";
                //        }
                //        else
                //        {
                //            OWNER_INFO = OWNER_INFO + ";12";
                //        }
                //    }
                //}
                for (int i = 0; i < PARCELTable.Rows.Count; i++)
                {
                    PARCELTable.Rows[i]["OWNER_INFO"] = "";
                    var card_no = OWNERTable.Select("PARCEL_CODE ='" + PARCELTable.Rows[i]["PARCEL_CODE"].ToString() + "'");
                    if (card_no.Length != 0)
                    {
                        int card_bool = 0;
                        for (int j = 0; j < card_no.Length; j++)
                        {
                            if (card_no[j]["CARD_NO"].ToString().Length > 5)
                            {
                                card_bool = 1;
                                break;
                            }
                        }
                        if (card_bool == 1)
                        {
                            PARCELTable.Rows[i]["OWNER_INFO"] = "12";
                        }
                    }
                    string parname = PARCELTable.Rows[i]["权利人"].ToString();
                    if (parname.Length != 0)
                    {
                        if (parname.Contains("业主"))
                        {
                            if (PARCELTable.Rows[i]["OWNER_INFO"].ToString() == "")
                            {
                                PARCELTable.Rows[i]["OWNER_INFO"] = "21";
                            }
                            else
                            {
                                PARCELTable.Rows[i]["OWNER_INFO"] = PARCELTable.Rows[i]["OWNER_INFO"].ToString() + ";21";
                            }

                        }
                    }
                    var code = PARCELTable2.Select("PARCEL_CODE = '" + PARCELTable.Rows[i]["PARCEL_CODE"].ToString() + "'");
                    if (code.Length != 0)
                    {
                        if (PARCELTable.Rows[i]["OWNER_INFO"].ToString() == "")
                        {
                            PARCELTable.Rows[i]["OWNER_INFO"] = "33";
                        }
                        else
                        {
                            PARCELTable.Rows[i]["OWNER_INFO"] = PARCELTable.Rows[i]["OWNER_INFO"].ToString() + ";33";
                        }

                    }
                }                
                List<string> sql1 = new List<string>();
                for (int i = 0; i < PARCELTable.Rows.Count; i++)
                {
                    string sqlname = "UPDATE LAND_CADASTRAL_SURVEY_PARCEL SET LAND_CADASTRAL_SURVEY_PARCEL.OWNER_INFO = '" +
                        PARCELTable.Rows[i]["OWNER_INFO"].ToString() + "' WHERE LAND_CADASTRAL_SURVEY_PARCEL.OBJECTID = " + Convert.ToInt32(PARCELTable.Rows[i]["OBJECTID"]);
                    sql1.Add(sqlname);
                }
                int res1 = OleDbHelper.RunTransAction(sql1, path);
                if (res1 == 0)
                {
                    System.Windows.Forms.MessageBox.Show("修改权利来源失败");
                }
                #endregion

                #region 修改权利人
                List<Owner> owners = new List<Owner>();
                for (int i = 0; i < PARCELTable.Rows.Count; i++)
                {
                    string ownername = PARCELTable.Rows[i]["权利人"].ToString();
                    if (ownername.Length != 0)
                    {
                        if (ownername.Contains(',') || ownername.Contains('，'))
                        {
                            string[] namearr = ownername.Split(new char[2] { ',', '，' });
                            int a = 1;
                            foreach (string name in namearr)
                            {
                                Owner ol = new Owner();
                                ol.PARCEL_CODE = PARCELTable.Rows[i]["PARCEL_CODE"].ToString();
                                ol.NAME = name;
                                ol.LCS_OWNER_ID = a++.ToString();
                                ol.PERSON_TYPE = "土地权利人";
                                ol.OWNED_TYPE = "1";
                                ol.OWNED_PORTION = "100";
                                owners.Add(ol);
                            }
                        }
                        else
                        {
                            Owner ol = new Owner();
                            ol.PARCEL_CODE = PARCELTable.Rows[i]["PARCEL_CODE"].ToString();
                            ol.NAME = PARCELTable.Rows[i]["权利人"].ToString();
                            ol.LCS_OWNER_ID = "1";
                            ol.PERSON_TYPE = "土地权利人";
                            ol.OWNED_TYPE = "0";
                            ol.OWNED_PORTION = "100";
                            owners.Add(ol);
                        }
                    }
                }

                for (int i = 0; i < owners.Count; i++)
                {
                    var row = OWNERTable.Select("NAME='" + owners[i].NAME + "'");
                    for (int j = 0; j < row.Length; j++)
                    {
                        if (row[j]["CARD_NO"].ToString() != "" || row[j]["CARD_NO"].ToString() != "/")
                        {
                            owners[i].CARD_NO = row[j]["CARD_NO"].ToString();
                        }
                        if (row[j]["OWNER_ADDRESS"].ToString() != "" || row[j]["OWNER_ADDRESS"].ToString() != "/")
                        {
                            owners[i].OWNER_ADDRESS = row[j]["OWNER_ADDRESS"].ToString();
                        }
                    }

                    if (owners[i].NAME.Length <= 3)
                    {
                        owners[i].TYPE = "1";
                        owners[i].CARD_TYPE = "居民身份证（居住证）";
                    }
                    if (owners[i].NAME.EndsWith("局") || owners[i].NAME.EndsWith("委员会") || owners[i].NAME.EndsWith("检察院"))
                    {
                        owners[i].TYPE = "2";
                        owners[i].CARD_TYPE = "统一社会信用代码证（机构代码证）";
                    }
                    if (owners[i].NAME.EndsWith("政府") || owners[i].NAME.EndsWith("街道办事处"))
                    {
                        owners[i].TYPE = "10";
                        owners[i].CARD_TYPE = "统一社会信用代码证（机构代码证）";
                    }
                    if (owners[i].NAME.EndsWith("公司") || owners[i].NAME.EndsWith("厂"))
                    {
                        owners[i].TYPE = "6";
                        owners[i].CARD_TYPE = "统一社会信用代码证（机构代码证）";
                    }
                    if (owners[i].NAME.EndsWith("业主"))
                    {
                        owners[i].TYPE = "99";
                        owners[i].CARD_TYPE = "居民身份证（居住证）";
                    }

                }

                List<string> sql = new List<string>();
                string ownerclear = "DELETE FROM LAND_CADASTRAL_SURVEY_OWNER";
                sql.Add(ownerclear);
                if(OWNERTable.Columns.Count == 14)
                {
                    for (int i = 0; i < owners.Count; i++)
                    {
                        string ownerinsert = "INSERT INTO LAND_CADASTRAL_SURVEY_OWNER VALUES" +
                            "('" + owners[i].LCS_OWNER_ID + "','" + owners[i].PARCEL_CODE + "','" + owners[i].PERSON_TYPE +
                            "','" + owners[i].NAME + "','" + owners[i].TYPE + "','" + owners[i].CARD_TYPE +
                            "','" + owners[i].CARD_NO + "','" + owners[i].OWNER_ADDRESS + "','" + owners[i].FRDBZM +
                            "','" + owners[i].OWNER_CONTACT + "','" + owners[i].OWNED_TYPE + "','" + owners[i].OWNED_PORTION +
                            "','" + owners[i].SURVEY_NO + "','" + owners[i].LCS_PARCEL_ID + "')";
                        sql.Add(ownerinsert);
                    }
                }
                else
                {
                    for (int i = 0; i < owners.Count; i++)
                    {
                        string ownerinsert = "INSERT INTO LAND_CADASTRAL_SURVEY_OWNER VALUES" +
                            "('" + owners[i].LCS_OWNER_ID + "','" + owners[i].PARCEL_CODE + "','" + owners[i].PERSON_TYPE +
                            "','" + owners[i].NAME + "','" + owners[i].TYPE + "','" + owners[i].CARD_TYPE +
                            "','" + owners[i].CARD_NO + "','" + owners[i].OWNER_ADDRESS + "','" + owners[i].FRDBZM +
                            "','" + owners[i].OWNER_CONTACT + "','" + owners[i].OWNED_TYPE + "','" + owners[i].OWNED_PORTION +
                            "')";
                        sql.Add(ownerinsert);
                    }
                }
                
                int res = OleDbHelper.RunTransAction(sql, path);
                if (res == 0)
                {
                    System.Windows.Forms.MessageBox.Show("修改权利人库失败");
                }
                #endregion

                #region 文档2
                List<string> sqllist = new List<string>();
                for (int i = 0; i < PARCELTable.Rows.Count; i++)
                {
                    var row = PARCELTable2.Select("PARCEL_NO = '" + PARCELTable.Rows[i]["宗地号_1"].ToString() + "'");
                    if (row.Length == 0)
                    {
                        continue;
                    }
                    var luarea = row[0]["LU_AREA"];
                    if (luarea.ToString().Length == 0)
                    {
                        luarea = null;
                    }
                    else
                    {
                        luarea = Convert.ToInt32(luarea);
                    }
                    var APPROVE_AREA = row[0]["LU_AREA"];
                    if (APPROVE_AREA.ToString().Length == 0)
                    {
                        APPROVE_AREA = null;
                    }
                    else
                    {
                        APPROVE_AREA = Convert.ToInt32(APPROVE_AREA);
                    }
                    if (row.Length != 0)
                    {                       
                        
                            string sql3 = "UPDATE LAND_CADASTRAL_SURVEY_PARCEL SET IS_REGISTER = '" + row[0]["IS_REGISTER"].ToString() + "', LU_FUNCTION = '" + row[0]["LU_FUNCTION"].ToString() +
                                "', OWNER_TYPE = '" + row[0]["OWNER_TYPE"].ToString() + "', OWNER_TYPE_CODE = '" + row[0]["OWNER_TYPE_CODE"].ToString() + "', OWNER_SOURCE_CODE = '" + row[0]["OWNER_SOURCE_CODE"].ToString() +
                                "', LU_AREA = " + luarea + ", APPROVE_NO = '" + row[0]["APPROVE_NO"].ToString() + "', PARCEL_TYPE_CODE = '" + row[0]["PARCEL_TYPE_CODE"].ToString() +
                                "', FILE_TYPE = '" + row[0]["FILE_TYPE"].ToString() + "', FILE_ID_OLD = '" + row[0]["FILE_ID_OLD"].ToString() + "', FILE_SUBTYPE = '" + row[0]["FILE_SUBTYPE"].ToString() +
                                "', APPROVE_AREA = " + APPROVE_AREA + ", PARCEL_CODE = '" + row[0]["PARCEL_CODE"].ToString() + "', LCS_QC_ID = '" + row[0]["LCS_QC_ID"].ToString() +
                                "' WHERE LAND_CADASTRAL_SURVEY_PARCEL.[宗地号_1] = '" + PARCELTable.Rows[i]["宗地号_1"].ToString() + "'";
                            sqllist.Add(sql3);
                        
                    }
                }
                int res2 = OleDbHelper.RunTransAction(sqllist, path);
                #endregion

                #region 文档3
                if (TextBox3_Copy.Text.Length != 0)
                {
                    string qs = "SELECT * FROM LAND_QSLYWJ_TB";
                    var qsTable = OleDbHelper.QueryTable(qs, path3);


                    List<string> sqllist1 = new List<string>();
                    for (int i = 0; i < JSTable.Rows.Count; i++)
                    {
                        var row = qsTable.Select("FILE_ID_OLD = '" + JSTable.Rows[i]["主键号"].ToString() + "'");
                        if (row.Length != 0)
                        {
                            string zjh = row[0]["FILE_ID_OLD"].ToString();

                            if (row[0]["LU_TERM"].ToString().Length == 0)
                            {
                                string sql2 = "UPDATE LAND_CADASTRAL_SURVEY_PARCEL SET LU_TERM = null, START_DATE =null, END_DATE =null WHERE LAND_CADASTRAL_SURVEY_PARCEL.[宗地号_1] = '" + PARCELTable.Rows[i]["宗地号_1"].ToString() + "'";
                                sqllist1.Add(sql2);
                            }
                            else if (row[0]["START_DATE"].ToString().Length == 0 || row[0]["END_DATE"].ToString().Length == 0)
                            {
                                string sql2 = "UPDATE LAND_CADASTRAL_SURVEY_PARCEL SET LU_TERM = '" + Convert.ToInt32(row[0]["LU_TERM"].ToString()) +
                                    "', START_DATE =null, END_DATE =null WHERE LAND_CADASTRAL_SURVEY_PARCEL.[宗地号_1] = '" + PARCELTable.Rows[i]["宗地号_1"].ToString() + "'";
                                sqllist1.Add(sql2);
                            }
                            else
                            {
                                string sql2 = "UPDATE LAND_CADASTRAL_SURVEY_PARCEL SET LU_TERM = '" +Convert.ToInt32(row[0]["LU_TERM"].ToString()) +
                                    "', START_DATE ='" + Convert.ToDateTime(row[0]["START_DATE"]) + "', END_DATE ='" + Convert.ToDateTime(row[0]["END_DATE"]) +
                                    "' WHERE LAND_CADASTRAL_SURVEY_PARCEL.[宗地号_1] = '" + PARCELTable.Rows[i]["宗地号_1"].ToString() + "'";
                                sqllist1.Add(sql2);
                            }

                        }
                    }
                    int res3 = OleDbHelper.RunTransAction(sqllist1, path);
                }
                #endregion


                System.Windows.Forms.MessageBox.Show("成功");

            }
            catch(Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("失败"+ ex);
            }
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            OpenFileDialog folderDlg = new OpenFileDialog();
            folderDlg.Filter = "ACCESS|*.mdb";
            folderDlg.ShowDialog();
            string filename = folderDlg.FileName;
            //FolderBrowserDialog.Description = "请选择数据库路径";
            //FolderBrowserDialog.ShowDialog();
            TextBox3_Copy.Text = "" + filename;

        }
    }
}
