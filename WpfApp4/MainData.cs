using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WpfApp4
{
    class MainData
    {
        string url_1 = "http://push2.eastmoney.com/api/qt/stock/get?ut=fa5fd1943c7b386f172d6893dbfba10b&invt=2&fltt=2&fields=f43,f57,f58,f169,f170,f46,f44,f51,f168,f47,f164,f163,f116,f60,f45,f52,f50,f48,f167,f117,f71,f161,f49,f530,f135,f136,f137,f138,f139,f141,f142,f144,f145,f147,f148,f140,f143,f146,f149,f55,f62,f162,f92,f173,f104,f105,f84,f85,f183,f184,f185,f186,f187,f188,f189,f190,f191,f192,f107,f111,f86,f177,f78,f110,f262,f263,f264,f267,f268,f250,f251,f252,f253,f254,f255,f256,f257,f258,f266,f269,f270,f271,f273,f274,f275,f127,f199,f128,f193,f196,f194,f195,f197,f80,f280,f281,f282,f284,f285,f286,f287,f292&secid=";
        //1.601818
        string url_2 = "&cb=jQuery112408937322727322846_1593942746483&_=1593942746484.js";

        //公司核心数据
        string Name;//公司名称 f58
        string Id;//股票代码 f57
        string ShouYi;//收益 f55
        string Pe;//PE（动）f162
        string MeiGuJingZiChan;//每股净资产 f92
        string ShiJingLv;//市净率 f167
        string ZongShouYi;//总收益（亿） f183
        string ZSY_TongBi;//总收益—同比（%） f184
        string JingLiRun;//净利润（亿） f105
        string JLR_TOngBi;//净利润-同比（%） f185
        string MaoLiLv;//毛利率（%） f186
        string JingLiLv;//净利率（%） f187
        string ROE;//（%） f173
        string FuZhaiLv;//负债率（%） f188
        string ZongGuBen;//总股本（亿） f84
        string ZongZhi;//总值（亿） f116
        string LiuTongGu;//流通股（亿） f85
        string LiuZhi;//流值（亿） f117
        string MeiGuWeiFenPeiLiRun;//每股未分配利润（元） f190

        public MainData(string id)
        {
            Id = id;
            Boolean r = GetData(Id);
            Console.WriteLine(r);
        }

        Boolean GetData(string id)
        {
            try
            {
                string url = url_1 + "0." + id + url_2;

                WebClient MyWebClient = new WebClient();
                MyWebClient.Credentials = CredentialCache.DefaultCredentials;
                Byte[] pageData = MyWebClient.DownloadData(url);
                string pageJson = Encoding.UTF8.GetString(pageData);
                if (pageJson.Length < 200)
                {
                    url = url_1 + "1." + id + url_2;
                    pageData = MyWebClient.DownloadData(url);
                    pageJson = Encoding.UTF8.GetString(pageData);
                }
                
                int index = pageJson.IndexOf("(");
                string json = pageJson.Substring(index + 1, pageJson.Length - index - 3);

                Newtonsoft.Json.Linq.JObject jObj = JObject.Parse(json);
                Console.WriteLine(json);
                JToken jName = jObj["data"]["f58"];
                Name = jName.ToString();
                JToken jShouYi = jObj["data"]["f55"];
                ShouYi = jShouYi.ToString();
                JToken jPe = jObj["data"]["f162"];
                Pe = jPe.ToString();
                JToken jMeiGuJingZiChan = jObj["data"]["f92"];
                MeiGuJingZiChan = jMeiGuJingZiChan.ToString();
                JToken jShiJingLv = jObj["data"]["f167"];
                ShiJingLv = jShiJingLv.ToString();
                JToken jZongShouYi = jObj["data"]["f183"];
                ZongShouYi = jZongShouYi.ToString();
                JToken jZSY_TongBi = jObj["data"]["f184"];
                ZSY_TongBi = jZSY_TongBi.ToString();
                JToken jJingLiRun = jObj["data"]["f105"];
                JingLiRun = jJingLiRun.ToString();
                JToken jJLR_TOngBi = jObj["data"]["f185"];
                JLR_TOngBi = jJLR_TOngBi.ToString();
                JToken jMaoLiLv = jObj["data"]["f186"];
                MaoLiLv = jMaoLiLv.ToString();
                JToken jJingLiLv = jObj["data"]["f187"];
                JingLiLv = jJingLiLv.ToString();
                JToken jROE = jObj["data"]["f173"];
                ROE = jROE.ToString();
                JToken jFuZhaiLv = jObj["data"]["f188"];
                FuZhaiLv = jFuZhaiLv.ToString();
                JToken jZongGuBen = jObj["data"]["f84"];
                ZongGuBen = jZongGuBen.ToString();
                JToken jZongZhi = jObj["data"]["f116"];
                ZongZhi = jZongZhi.ToString();
                JToken jLiuTongGu = jObj["data"]["f85"];
                LiuTongGu = jLiuTongGu.ToString();
                JToken jLiuZhi = jObj["data"]["f117"];
                LiuZhi = jLiuZhi.ToString();
                JToken jMeiGuWeiFenPeiLiRun = jObj["data"]["f190"];
                MeiGuWeiFenPeiLiRun = jMeiGuWeiFenPeiLiRun.ToString();
            }
            catch(Exception e)
            {
                return false;
            }
            return true;
        }

        public string[] GetData()
        {
            string[] data = { Name, Id, ShouYi, Pe, MeiGuJingZiChan, ShiJingLv, ZongShouYi,
                                ZSY_TongBi, JingLiRun, JLR_TOngBi, MaoLiLv, JingLiLv,
                                ROE, FuZhaiLv, ZongGuBen, ZongZhi, LiuTongGu, LiuZhi, MeiGuWeiFenPeiLiRun};

            return data;
        }

    }


}
