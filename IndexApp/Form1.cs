using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Text;
using System.Windows.Forms;
using Trade2015;
using Quote2015;
using System.Net;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Text.RegularExpressions;

namespace IndexApp
{
    public partial class Form1 : Form
    {
        //internal Trade trade;
        internal Quote quote;
        //private string[] indexesId = { "SR", "ru", "TA", "p", "zn", "l", "CF", "c", "RI", "WH", "v", "a", "OI", "al", "au", "y", "cu", "m", "rb", "pb", "j", "MA", "ag", "IF", "RM", "FG", "bu", "TC", "i", "jd", "fb", "bb", "TF", "jm", "pp", "hc" };
        //private string[] indexesId = { "SR", "RU", "TA", "P", "ZN", "L", "CF", "C", "RI", "WH", "V", "A", "OI", "AL", "AU", "Y", "CU", "M", "RB", "PB", "J", "MA", "AG", "IF", "RM", "FG", "BU", "TC", "I", "JD", "FB", "BB", "TF", "JM", "PP", "HC" };
        //private string[] indexesName = { "糖", "橡胶", "PTA", "棕榈油", "锌", "塑料", "棉花", "玉米", "早稻", "郑麦", "PVC", "豆一", "郑油", "铝", "黄金", "豆油", "铜", "豆粕", "螺纹", "铅", "焦炭", "郑醇", "白银", "期指", "菜粕", "玻璃", "沥青", "动煤", "铁矿", "鸡蛋", "纤维板", "胶合板", "国债", "焦煤", "PP", "热卷" };
        private string[] indexesId = { "CU", "SR", "RB", "Y", "M", "AL", "RU", "P", "ZN", "CF", "A", "RM", "AU", "FG", "TA", "L", "C", "OI", "RI", "V", "IF", "WH", "PB", "J", "MA", "AG", "BU", "TC", "I", "JD", "FB", "BB", "TF", "JM", "PP", "HC" };

        private string[] indexesName = { "铜", "白糖", "螺纹", "豆油", "豆粕", "铝", "橡胶", "棕榈油", "锌", "棉花", "大豆1", "菜粕", "黄金", "玻璃", "PTA", "塑料", "玉米", "郑油", "早稻", "PVC", "期指", "郑麦", "铅", "焦炭", "甲醇", "白银", "沥青", "动煤", "铁矿", "鸡蛋", "纤维板", "胶合板", "国债", "焦煤", "PP", "热卷" };

        //指数table
        //private DataTable dt = new DataTable("指数");
        private string colName;
        private int mainDtIndex = 0;
        //指数及各合约table list
        private List<DataTable> dtList = new List<DataTable>();

        private DataGridView settingGridView = new DataGridView();


        //设置table
        private DataTable settingDT = new DataTable();
        //public int[] tradeUnit = { 5, 10, 10, 5, 5, 5, 10, 5, 5, 5, 5, 10, 1000, 20, 5, 5, 5, 6, 5, 5, 600, 6, 5, 100, 10, 15, 10, 200, 100, 10, 500, 500, 20000, 60, 5, 10 };
        //public double[] tradeMargin = { 7, 6, 6, 10, 10, 5, 8, 5, 6, 5, 10, 5, 6, 6, 6, 5, 10, 10, 20, 5, 8, 20, 5, 7, 5, 8, 7, 5, 7, 8, 7, 10, 2, 7, 5, 6 };
        List<int> tradeUnit = new List<int>();
        List<double> tradeMargin = new List<double>();
        public Form1()
        {
            InitializeComponent();
        }

        private void connectBt_Click(object sender, EventArgs e)
        {
            this.connectBt.Text = "登录中...";

            quote = new Quote("ctp_quote_proxy.dll")
            {
                Server = "tcp://ctp1-md9.citicsf.com:41213",
                Broker = "66666",
                //Investor = "00090",
                //Password = "888888",
            };
            quote.OnFrontConnected += quote_OnFrontConnected;
            quote.OnRspUserLogin += quote_OnRspUserLogin;
            quote.ReqConnect();

            //trade = new Trade("ctp_trade_proxy.dll")
            //{
            //    //Server = "tcp://180.166.165.179:41205",
            //    //Broker = "1013",
            //    //Investor = "00000004",
            //    //Password = "123456",

            //    Server = "tcp://123.233.249.154:41205",
            //    Broker = "1040",
            //    Investor = "00000003",
            //    Password = "123456",
            //};
            //trade.OnFrontConnected += trade_OnFrontConnected;
            //trade.OnRspUserLogin += trade_OnRspUserLogin;
            //trade.ReqConnect();
        }

        void quote_OnFrontConnected(object sender, EventArgs e)
        {
            ((Quote)sender).ReqUserLogin();
        }

        void quote_OnRspUserLogin(object sender, Quote2015.IntEventArgs e)
        {
            LoginSuccess();
        }

        void LoginSuccess()
        {
            //trade.OnFrontConnected -= trade_OnFrontConnected;
            //trade.OnRspUserLogin -= trade_OnRspUserLogin;
            if (quote != null)
            {
                quote.OnFrontConnected -= quote_OnFrontConnected;
                quote.OnRspUserLogin -= quote_OnRspUserLogin;
                this.Invoke(new Action(() =>
                {
                    this.connectBt.Text = "服务器已连接";
                    this.getIdBt.Enabled = true;
                }));
            }
        }

        //void trade_OnFrontConnected(object sender, EventArgs e)
        //{
        //    ((Trade)sender).ReqUserLogin();
        //}

        //void trade_OnRspUserLogin(object sender, Trade2015.IntEventArgs e)
        //{
        //    if (e.Value == 0)
        //    {

        //        this.Invoke(new Action(() =>
        //        {
        //            this.connectBt.Text = "服务器已连接";
        //            this.getIdBt.Enabled = true;
        //        }));
        //        if (quote == null)
        //            LoginSuccess();
        //        else
        //            quote.ReqConnect();

        //    }
        //    else
        //    {
        //        this.Invoke(new Action(() =>
        //        {
        //            this.connectBt.Text = "连接服务器";
        //        }));
        //        MessageBox.Show("连接服务器失败");
        //        //ShowMsg("登录错误:" + e.Value.ToString());
        //        trade.ReqUserLogout();
        //        trade = null;
        //        //quote = null;
        //    }
        //}

        private void getIdBt_Click(object sender, EventArgs e)
        {
            DataTable mainDt = new DataTable("指数");
            dtList.Add(mainDt);
            if (dtList.Contains(mainDt))
            {
                mainDtIndex = dtList.IndexOf(mainDt);
                //dt第一列

                dtList[mainDtIndex].Columns.Add("品种", System.Type.GetType("System.String"));
                //dt第二列 均价
                dtList[mainDtIndex].Columns.Add("指数均价", System.Type.GetType("System.Double"));
                //dt第三列 持仓量
                //colName = "持仓量(" + DateTime.Now.ToString() + ")";
                colName = "持仓量";
                dtList[mainDtIndex].Columns.Add(colName, System.Type.GetType("System.Double"));

                dtList[mainDtIndex].Columns.Add("保证金", System.Type.GetType("System.Double"));

                //获取合约均价和持仓量，计算指数均价
                catchAndCal(indexesId, indexesName);
                //catchAndCalFromCtp(indexesId, indexesName);
                foreach (DataTable dt in dtList)
                {
                    //创建DataGridView并链接Datatable
                    DataGridView dgv = new DataGridView();
                    dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                    dgv.Margin = new Padding(2, 2, 2, 2);
                    dgv.Name = "dataGridView" + dt.TableName;
                    dgv.Dock = DockStyle.Fill;
                    dgv.RowHeadersVisible = false;
                    dgv.DataSource = dt;
                    //把Datagridview放入Tabpage
                    TabPage tp = new TabPage();
                    tp.Controls.Add(dgv);
                    tp.Text = dt.TableName;
                    tp.Name = "tabPage" + dt.TableName;
                    tp.UseVisualStyleBackColor = true;
                    //把Tabpage放入Tabcontrol
                    this.tabControl1.Controls.Add(tp);

                }
                this.exportBt.Enabled = true;
            }

        }

        private void catchAndCal(string[] indId, string[] indName)
        {
            this.mainTxtBox.Text = "";
            if (indId.Length == indName.Length)
            {
                List<string> allId = getAllID();
                //string[] allId = trade.DicInstrumentField.Keys.ToArray();

                this.mainTxtBox.Text += DateTime.Now.ToString("H:mm:ss") + "|| 合约名称" + allId.Count.ToString() + "个" + System.Environment.NewLine;
                for (int j = 0; j < indId.Length; j++)
                {
                    List<string> ids = new List<string>();
                    for (int i = 0; i < allId.Count; i++)
                    {
                        if (allId[i].Contains(indId[j]))
                        {
                            ids.Add(allId[i]);
                        }
                    }
                    if (ids.Count() != 0)
                    {
                        //创建并配置品种datatable
                        DataTable dt = new DataTable(indName[j]);
                        dtList.Add(dt);
                        int dtIndex = 0;
                        dtIndex = dtList.IndexOf(dt);
                        dtList[dtIndex].Columns.Add("合约", System.Type.GetType("System.String"));
                        dtList[dtIndex].Columns.Add("均价", System.Type.GetType("System.Double"));
                        dtList[dtIndex].Columns.Add("持仓量", System.Type.GetType("System.Double"));


                        List<double> settlementPrice = new List<double>();
                        List<int> openInterest = new List<int>();
                        double indexSettlementPrice;
                        idDuplicate(indId[j], ids);
                        ids.Sort();
                        if (indId[j] == "IF" || indId[j] == "TF")
                        {
                            foreach (string _id in ids)
                            {
                                quote.ReqSubscribeMarketData(_id);
                            }
                            Thread.Sleep(1500);
                        }
                        if (ids.Count > 12)
                        {
                            ids.RemoveRange(12, ids.Count - 12);
                        }
                        foreach (string id in ids)
                        {
                            //依次将合约信息保存在dt里
                            DataRow dr = dtList[dtIndex].NewRow();
                            dr["合约"] = id;
                            MarketData tick;
                            if (id.Contains("IF") || id.Contains("TF"))
                            {
                                this.mainTxtBox.Text += id + " from CTP" + System.Environment.NewLine;
                                if (quote.DicTick.TryGetValue(id, out tick))
                                {
                                    dr["均价"] = tick.SettlementPrice;
                                    dr["持仓量"] = tick.OpenInterest;
                                    settlementPrice.Add(tick.SettlementPrice);
                                    openInterest.Add(Convert.ToInt32(tick.OpenInterest));
                                }
                            }
                            else
                            {
                                this.mainTxtBox.Text += id + " from Web" + System.Environment.NewLine;
                                dr["均价"] = getSettlementPrice(id);
                                settlementPrice.Add(getSettlementPrice(id));
                                dr["持仓量"] = Convert.ToDouble(getOpenInterest(id));
                                openInterest.Add(getOpenInterest(id));
                            }
                            dtList[dtIndex].Rows.Add(dr);
                        }
                        indexSettlementPrice = getIndex(settlementPrice, openInterest);
                        //向dt中写入指数及均价
                        DataRow indexDr = dtList[mainDtIndex].NewRow();
                        indexDr["品种"] = indName[j];
                        indexDr["指数均价"] = Convert.ToDouble(priceWithType(indexesId[j], indexSettlementPrice));
                        indexDr[colName] = Convert.ToDouble(getIndexOI(openInterest));
                        //计算保证金
                        double indexSP = Convert.ToDouble(priceWithType(indexesId[j], indexSettlementPrice));
                        double indexOpenI = Convert.ToDouble(getIndexOI(openInterest));
                        double marginRate = tradeMargin[j] / 100;
                        double indexMargin = indexSP * indexOpenI * tradeUnit[j] * marginRate;
                        indexDr["保证金"] = Convert.ToDouble(indexMargin.ToString("0.0"));
                        dtList[mainDtIndex].Rows.Add(indexDr);

                        this.mainTxtBox.Text += DateTime.Now.ToString("H:mm:ss") + "|| " + indexesName[j] + "指数 " + "√" + System.Environment.NewLine;
                        
                    }

                }
            }
        }

        private void catchAndCalFromCtp(string[] indId, string[] indName)
        {
            this.mainTxtBox.Text = "";
            if (indId.Length == indName.Length)
            {
                List<string> allId = getAllID();
                //string[] allId = trade.DicInstrumentField.Keys.ToArray();


                this.mainTxtBox.Text += DateTime.Now.ToString("H:mm:ss") + "|| 合约共有" + allId.Count.ToString() + "个" + System.Environment.NewLine;
                for (int j = 0; j < indId.Length; j++)
                {
                    List<string> ids = new List<string>();
                    for (int i = 0; i < allId.Count; i++)
                    {
                        if (allId[i].Contains(indId[j]))
                        {
                            ids.Add(allId[i]);
                            //quote.ReqSubscribeMarketData(allId[i]);
                        }
                    }
                    if (ids.Count() != 0)
                    {
                        //创建并配置品种datatable
                        DataTable dt = new DataTable(indName[j]);
                        dtList.Add(dt);
                        int dtIndex = 0;
                        dtIndex = dtList.IndexOf(dt);
                        dtList[dtIndex].Columns.Add("合约", System.Type.GetType("System.String"));
                        dtList[dtIndex].Columns.Add("昨结算价", System.Type.GetType("System.Double"));
                        dtList[dtIndex].Columns.Add("昨持仓量", System.Type.GetType("System.Double"));


                        List<double> settlementPrice = new List<double>();
                        List<int> openInterest = new List<int>();
                        double indexSettlementPrice;

                        idDuplicate(indId[j], ids);
                        ids.Sort();
                        foreach (string _id in ids)
                        {
                            quote.ReqSubscribeMarketData(_id);
                        }
                        Thread.Sleep(2000);
                        foreach (string id in ids)
                        {
                            string newid = id;
                            if (id.Contains("SR") || id.Contains("CF") || id.Contains("FG") || id.Contains("OI") || id.Contains("RI") || id.Contains("TC") || id.Contains("TA") || id.Contains("WH"))
                            {
                                newid = id.Insert(2, "1");
                            }
                            else if (id.Contains("IF"))
                            {
                                newid = id.Insert(0, "CFF_");
                            }
                            MarketData tick;
                            
                            if (quote.DicTick.TryGetValue(id, out tick))
                            {
                                //依次将合约信息保存在dt里
                                DataRow dr = dtList[dtIndex].NewRow();
                                dr["合约"] = id;
                                this.mainTxtBox.Text += "get " + id + " from CTP" + System.Environment.NewLine;
                                dr["昨结算价"] = tick.SettlementPrice;
                                dr["昨持仓量"] = tick.OpenInterest;
                                settlementPrice.Add(tick.SettlementPrice);
                                openInterest.Add(Convert.ToInt32(tick.OpenInterest));
                                dtList[dtIndex].Rows.Add(dr);
                            }
                        }
                        indexSettlementPrice = getIndex(settlementPrice, openInterest);
                        //向dt中写入指数及均价
                        DataRow indexDr = dtList[mainDtIndex].NewRow();
                        indexDr["品种"] = indName[j];
                        indexDr["指数昨结"] = Convert.ToDouble(priceWithType(indexesId[j], indexSettlementPrice));
                        indexDr[colName] = Convert.ToDouble(getIndexOI(openInterest));
                        dtList[mainDtIndex].Rows.Add(indexDr);
                        this.mainTxtBox.Text += DateTime.Now.ToString("H:mm:ss") + "|| " + indexesName[j] + "指数 " + "√" + System.Environment.NewLine;


                        //this.Invoke(new Action(() =>
                        //{
                        //    this.mainTxtBox.Text += DateTime.Now.ToString("H:mm:ss") + "|| " + indexesName[j] + "指数 " + "√" + System.Environment.NewLine;
                        //}));

                    }

                }
            }
        }


        private List<string> getAllID()
        {
            List<string> allId = new List<string>();
            string url = "http://money.finance.sina.com.cn/d/api/openapi_proxy.php/?__s=[[%22qhhq%22,%22qbhy%22,%22zdf%22,1000]]&callback=getData.futures_qhhq_gnqh";
            string strMsg = string.Empty;
            try
            {
                WebRequest request = WebRequest.Create(url);
                WebResponse response = request.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.GetEncoding("gb2312"));
                strMsg = reader.ReadToEnd();

                reader.Close();
                reader.Dispose();
                response.Close();
            }
            catch
            { }
            string pattern = @"\[\[";
            Regex regex = new Regex(pattern);
            string[] strPart = regex.Split(strMsg);
            if (strPart.Length >= 2)
            {
                string[] subStr = strPart[1].Split('[');
                foreach (string str in subStr)
                {
                    string[] finalStr = str.Split(',');
                    if (finalStr.Length > 2)
                    {
                        string idStr = finalStr[1].Trim('"');
                        allId.Add(idStr);
                    }
                }
                //List<string> idAfterAjust = idAjust(allId);
                //return idAfterAjust;
                return allId;
            }
            else
            {
                MessageBox.Show("获取合约失败，请联系金海天");
                return allId;
            }

        }

        private List<string> idAjust(List<string> originalId)
        {
            List<string> newAllId = new List<string>();
            foreach (string id in originalId)
            {
                string newId = id;
                if (id.Contains("SR") || id.Contains("TA") || id.Contains("CF") || id.Contains("RI") || id.Contains("WH") || id.Contains("OI") || id.Contains("FG") || id.Contains("TC") || id.Contains("MA") || id.Contains("RM"))
                {
                    if (id.Length > 3)
                    {
                        newId = id.Remove(2, 1);
                    }
                }
                else if (id.Contains("IF") || id.Contains("TF"))
                {

                }
                else
                {
                    newId = id.ToLower();
                }
                newAllId.Add(newId);
            }
            return newAllId;
        }

        private void idDuplicate(string indexType, List<string> idList)
        {
            //玉米合约检索时会混入铜合约，所以要除去铜合约
            if (indexType.Equals("C"))
            {
                List<string> removeId = new List<string>();
                for (int k = 0; k < idList.Count(); k++)
                {
                    if (idList[k].Contains("CU") || idList[k].Contains("CS") || idList[k].Contains("HC") || idList[k].Contains("CF") || idList[k].Contains("TC"))
                    {
                        removeId.Add(idList[k]);
                    }
                }
                foreach (string id in removeId)
                {
                    idList.Remove(id);
                }
            }
            //豆一理由同上
            if (indexType.Equals("A"))
            {

                List<string> removeId = new List<string>();
                for (int k = 0; k < idList.Count(); k++)
                {
                    if (idList[k].Contains("AU") || idList[k].Contains("AL") || idList[k].Contains("AG") || idList[k].Contains("MA") || idList[k].Contains("TA"))
                    {
                        removeId.Add(idList[k]);
                    }
                }
                foreach (string id in removeId)
                {
                    idList.Remove(id);
                }
            }
            //棕榈同上
            if (indexType.Equals("P"))
            {
                List<string> removeId = new List<string>();
                for (int k = 0; k < idList.Count(); k++)
                {
                    if (idList[k].Contains("PB") || idList[k].Contains("PP"))
                    {
                        removeId.Add(idList[k]);
                    }
                }
                foreach (string id in removeId)
                {
                    idList.Remove(id);
                }
            }
            if (indexType.Equals("L"))
            {
                List<string> removeId = new List<string>();
                for (int k = 0; k < idList.Count(); k++)
                {
                    if (idList[k].Contains("AL") || idList[k].Contains("LR"))
                    {
                        removeId.Add(idList[k]);
                    }
                }
                foreach (string id in removeId)
                {
                    idList.Remove(id);
                }
            }

            if (indexType.Equals("M"))
            {
                List<string> removeId = new List<string>();
                for (int k = 0; k < idList.Count(); k++)
                {
                    if (idList[k].Contains("JM") || idList[k].Contains("SM") || idList[k].Contains("MA") || idList[k].Contains("RM"))
                    {
                        removeId.Add(idList[k]);
                    }
                }
                foreach (string id in removeId)
                {
                    idList.Remove(id);
                }
            }
            if (indexType.Equals("J"))
            {
                List<string> removeId = new List<string>();
                for (int k = 0; k < idList.Count(); k++)
                {
                    if (idList[k].Contains("JD") || idList[k].Contains("JM") || idList[k].Contains("JR"))
                    {
                        removeId.Add(idList[k]);
                    }
                }
                foreach (string id in removeId)
                {
                    idList.Remove(id);
                }
            }
            if (indexType.Equals("I"))
            {
                List<string> removeId = new List<string>();
                for (int k = 0; k < idList.Count(); k++)
                {
                    if (idList[k].Contains("IF") || idList[k].Contains("OI") || idList[k].Contains("RI"))
                    {
                        removeId.Add(idList[k]);
                    }
                }
                foreach (string id in removeId)
                {
                    idList.Remove(id);
                }
            }
        }

        //取得结算价(double)
        private double getSettlementPrice(string quoteId)
        {
            string url = "http://hq.sinajs.cn/list=" + quoteId.ToUpper();
            double settlementPrice = 0;
            string strMsg = string.Empty;
            try
            {
                WebRequest request = WebRequest.Create(url);
                WebResponse response = request.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.GetEncoding("gb2312"));
                strMsg = reader.ReadToEnd();

                reader.Close();
                reader.Dispose();
                response.Close();
            }
            catch
            { }
            string[] strPart = strMsg.Split(',');
            if (strPart.Length > 14)
            {
                settlementPrice = Convert.ToDouble(strPart[9]);
            }
            return settlementPrice;
        }


        //取得持仓量(int)
        private int getOpenInterest(string quoteId)
        {
            string url = "http://hq.sinajs.cn/list=" + quoteId.ToUpper();
            int openInterest = 0;
            string strMsg = string.Empty;
            try
            {
                WebRequest request = WebRequest.Create(url);
                WebResponse response = request.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.GetEncoding("gb2312"));

                strMsg = reader.ReadToEnd();

                reader.Close();
                reader.Dispose();
                response.Close();
            }
            catch
            { }
            string[] strPart = strMsg.Split(',');
            if (strPart.Length > 14)
            {
                decimal tmp = Convert.ToDecimal(strPart[13]);
                openInterest = Convert.ToInt32(tmp);
            }
            return openInterest;
        }

        //计算指数均价
        private double getIndex(List<double> SPList, List<int> OIList)
        {
            double index = 0;
            if (SPList.Count() == OIList.Count())
            {
                double dotProduct = 0;
                int oiSum = 0;
                for (int i = 0; i < SPList.Count(); i++)
                {
                    dotProduct += SPList[i] * OIList[i];
                    oiSum += OIList[i];
                }
                if (oiSum != 0)
                {
                    index = Convert.ToDouble(dotProduct / oiSum);
                }
            }
            return index;
        }

        //计算指数持仓量
        private int getIndexOI(List<int> OIList)
        {
            int indexOI = 0;
            //foreach (int oi in OIList)
            //{
            //    indexOI += oi;
            //}
            for (int i = 0; i < OIList.Count(); i++)
            {
                indexOI += OIList[i];
            }
            return indexOI;
        }

        private string priceWithType(string indexId, double stPrice)
        {
            if (indexId.Contains("TF"))
            {
                return stPrice.ToString("0.000");
            }
            else if (indexId.Contains("IF") || indexId.Contains("TC"))
            {
                return stPrice.ToString("0.0");
            }
            else if (indexId.Contains("au") || indexId.Contains("bb") || indexId.Contains("fb"))
            {
                return stPrice.ToString("0.00");
            }
            else
            {
                //return Convert.ToInt32(stPrice).ToString();
                return stPrice.ToString("0");
            }
        }

        public static void TableToExcelForXLSX(List<DataTable> listDT, string file)
        {
            XSSFWorkbook xssfworkbook = new XSSFWorkbook();
            foreach (DataTable dt in listDT)
            {
                ISheet sheet = xssfworkbook.CreateSheet(dt.TableName);
                //表头  
                IRow row = sheet.CreateRow(0);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    ICell cell = row.CreateCell(i);
                    cell.SetCellValue(dt.Columns[i].ColumnName);
                }

                //数据
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    IRow row1 = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ICell cell = row1.CreateCell(j);
                        if (dt.Rows[i][j].GetType().ToString().Equals("System.Double"))
                        {
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(Convert.ToDouble(dt.Rows[i][j]));
                        }
                        else
                        {
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                }

            }

            //转为字节数组  
            MemoryStream stream = new MemoryStream();
            xssfworkbook.Write(stream);
            var buf = stream.ToArray();

            //保存为Excel文件  
            using (FileStream fs = new FileStream(file, FileMode.OpenOrCreate, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
        }

        //public static void UpdateExcelForXLSX(List<DataTable> listDT, string file)
        //{
        //    FileStream fs = new FileStream(file, FileMode.Open, FileAccess.ReadWrite);
        //    XSSFWorkbook xssfworkbook = new XSSFWorkbook(fs);
        //    for (int i = 0; i < listDT.Count; i++)
        //    {
        //        ISheet sheet = xssfworkbook.GetSheetAt(i);

        //        //数据
        //        for (int j = 0; j < listDT[i].Rows.Count; j++)
        //        {
        //            IRow row1 = sheet.GetRow(j + 1);
        //            for (int k = 0; k < listDT[i].Columns.Count; k++)
        //            {
        //                ICell cell = row1.GetCell(k);
        //                if (listDT[i].Rows[j][k].GetType().ToString().Equals("System.Double"))
        //                {
        //                    cell.SetCellType(CellType.Numeric);
        //                    cell.SetCellValue(Convert.ToDouble(listDT[i].Rows[j][k]));
        //                }
        //                else
        //                {
        //                    cell.SetCellValue(listDT[i].Rows[j][k].ToString());
        //                }
        //            }
        //        }
        //    }
        //
        //    //转为字节数组  
        //    MemoryStream stream = new MemoryStream();
        //    xssfworkbook.Write(stream);
        //    var buf = stream.ToArray();
        //    using (FileStream fs1 = new FileStream(file, FileMode.OpenOrCreate, FileAccess.Write))
        //    {
        //        fs1.Write(buf, 0, buf.Length);
        //        fs1.Flush();
        //    }
        //}

        private void Form1_Load(object sender, EventArgs e)
        {
            this.getIdBt.Enabled = false;
            this.exportBt.Enabled = false;
            if (File.Exists(Application.StartupPath + "\\tradeunit.txt"))
            {
                string str;
                StreamReader sr = new StreamReader(Application.StartupPath + "\\tradeunit.txt", false);
                str = sr.ReadLine().ToString();
                sr.Close();
                string[] tradeUnitFromTxt = str.Split(',');
                foreach (string tuft in tradeUnitFromTxt)
                {
                    tradeUnit.Add(Convert.ToInt32(tuft));
                }
            }
            else
            {
                MessageBox.Show("交易单位文件丢失，请联系金海天");
            }
            if (File.Exists(Application.StartupPath + "\\trademargin.txt"))
            {
                string str;
                StreamReader sr = new StreamReader(Application.StartupPath + "\\trademargin.txt", false);
                str = sr.ReadLine().ToString();
                sr.Close();
                string[] tradeMarginFromTxt = str.Split(',');
                foreach (string tmft in tradeMarginFromTxt)
                {
                    tradeMargin.Add(Convert.ToDouble(tmft));
                }
            }
            else
            {
                MessageBox.Show("保证金文件丢失，请联系金海天");
            }
            



            //setting
            settingDT.Columns.Add("品种");
            settingDT.Columns.Add("交易单位");
            settingDT.Columns.Add("保证金(%)");
            if (indexesName.Length == tradeUnit.Count && indexesName.Length == tradeMargin.Count)
            {
                for (int i = 0; i < indexesName.Length; i++)
                {
                    DataRow dr = settingDT.NewRow();
                    dr["品种"] = indexesName[i];
                    dr["交易单位"] = tradeUnit[i];
                    dr["保证金(%)"] = tradeMargin[i];
                    settingDT.Rows.Add(dr);
                }

                
                settingGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                settingGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                settingGridView.Margin = new Padding(0, 0, 0, 0);
                settingGridView.Name = "settingGridView";
                settingGridView.Dock = DockStyle.Fill;
                settingGridView.RowHeadersVisible = false;
                settingGridView.DataSource = settingDT;
                settingGridView.CellValueChanged += new DataGridViewCellEventHandler(this.CellValueChanged);
                

                settingTabPage.Controls.Add(settingGridView);
                settingTabPage.Padding = new Padding(0, 0, 0, 0);
                //tp.UseVisualStyleBackColor = true;

            }
        }

        private void exportBt_Click(object sender, EventArgs e)
        {
            //excel文件地址
            string file = this.excelPathTxtBox.Text + "\\指数数据" + ".xlsx";
            //将dt数据写入excel文件

            if (File.Exists(file))
            {
                File.Delete(file);
            }
            TableToExcelForXLSX(dtList, file);
            //Thread.Sleep(5000);
            //if (File.Exists(file))
            //{
            //    System.Diagnostics.Process.Start(file);
            //}

        }

        private void folderSelectBt_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();
            if (result == DialogResult.OK)
            {
                this.excelPathTxtBox.Text = fbd.SelectedPath;
            }
        }

        private void CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (indexesName.Length == tradeUnit.Count && indexesName.Length == tradeMargin.Count)
            {
                if (e.ColumnIndex == 1)
                {
                    string tradeUnitToTxt = "";
                    tradeUnit[e.RowIndex] =Convert.ToInt32(this.settingGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
                    for (int i = 0; i < tradeUnit.Count - 1; i++)
                    {
                        tradeUnitToTxt += tradeUnit[i].ToString() + ",";
                    }
                    tradeUnitToTxt += tradeUnit[tradeUnit.Count - 1];
                    StreamWriter sw = new StreamWriter(Application.StartupPath +"\\tradeunit.txt", false);
                    sw.WriteLine(tradeUnitToTxt);
                    sw.Close();
                }
                else if (e.ColumnIndex == 2)
                {
                    string tradeMarginToTxt = "";
                    tradeMargin[e.RowIndex] = Convert.ToDouble(this.settingGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
                    for (int i = 0; i < tradeMargin.Count - 1; i++)
                    {
                        tradeMarginToTxt += tradeMargin[i].ToString() + ",";
                    }
                    tradeMarginToTxt += tradeMargin[tradeMargin.Count - 1];
                    StreamWriter sw = new StreamWriter(Application.StartupPath + "\\trademargin.txt", false);
                    sw.WriteLine(tradeMarginToTxt);
                    sw.Close();
                }
            }
            
        }

        


    }
}
