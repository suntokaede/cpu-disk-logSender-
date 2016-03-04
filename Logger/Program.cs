using System;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics.Eventing.Reader;

namespace Logger {

    class Program {

        static void Main() {
            Run_Async().Wait();
        }

        static async Task Run_Async() {
            string APIKey = GetConfig("APIKey");
            string RoomId = GetConfig("RoomId");

            //PerformanceCounterの準備
            PerformanceCounter pc_cpu = new PerformanceCounter("Processor", "% Processor Time", "_Total", true);
            PerformanceCounter pc_disk = new PerformanceCounter("LogicalDisk", "% Free Space", "_Total", true);

            float cpuUsageThreshold, diskUsageThreshold;

            //pc_cpu.NextValue()の最初の呼び出しでは何故か0が返ってくるので、一回NextValueを呼び出しています。
            float firstValue = pc_cpu.NextValue();

            //App.configから閾値の呼び出し
            if (!float.TryParse(GetConfig("CPUUsageThreshold"), out cpuUsageThreshold) ||
                !float.TryParse(GetConfig("DiskUsageThreshold"), out diskUsageThreshold)) {
                throw new System.ArgumentException("閾値を読み込むことができませんでした。");
            }

            //閾値が0～1を超えていないかチェック
            if (cpuUsageThreshold < 0 || cpuUsageThreshold > 1.0 || diskUsageThreshold < 0 || diskUsageThreshold > 1.0) {
                throw new System.ArgumentOutOfRangeException("閾値が無効です。");
            }

            //CPU使用率は、一回NextValueを実行しただけでは正確な値を取得できないので、１０回ループを回してその平均値を取っています。
            float cpuUsageSum = 0;
            for (int i = 0; i < 10; i++) {
                cpuUsageSum += pc_cpu.NextValue();
                await Task.Delay(500);
            }
            double cpuUsage = cpuUsageSum / 1000;
            float diskUsage = (100 - pc_disk.NextValue()) / 100;

            //flag判定
            WarningFlags flags = WarningFlags.None;
            if (cpuUsage >= cpuUsageThreshold) { flags |= WarningFlags.Cpu; }
            if (diskUsage >= diskUsageThreshold) { flags |= WarningFlags.Disk; }

            if (flags > 0) {

                StringBuilder bldr = new StringBuilder();
                if ((flags & WarningFlags.Cpu) == WarningFlags.Cpu) {
                    bldr.AppendFormat("CPU Usage : {0:p0} >= {1:p0}\r\n", cpuUsage, cpuUsageThreshold);
                }
                if ((flags & WarningFlags.Disk) == WarningFlags.Disk) {
                    bldr.AppendFormat("Disk Usage : {0:p0} >= {1:p0}\r\n", diskUsage, diskUsageThreshold);
                }

                //イベントログ取得処理
                var query = new EventLogQuery("System", PathType.LogName, getQuery());
                using (var reader = new EventLogReader(query)) {
                    // 直近 100 件のイベントレコードを表示
                    List<EventRecord> records = reader.ReadAllEvents()
                                                             .Reverse()
                                                             .Take(100)
                                                             .ToList();

                    foreach (EventRecord record in records) {
                        bldr.AppendFormat("[{0}] {1:yyyy/MM/dd HH:mm} {2} {3} {4}\r\n", record.LevelDisplayName, record.TimeCreated, record.LogName, record.ProviderName, record.Id);
                    }

                    //最も新しいレコードの日付をセットします。
                    if (records.Any()) {
                        LastRunTime = (DateTime)records.First().TimeCreated;
                    }
                }
                string info = "CPU, Disk and Log Info at " + DateTime.Now;
                SendMessage(APIKey, RoomId, info, bldr.ToString()).Wait();

            }

        }

        static DateTime LastRunTime {
            get {
                //プロパティが空の場合は直近10日分のDateTimeを返します
                if (Properties.Settings.Default["LastDate"] == null) {
                    return DateTime.Now.AddDays(-10);
                }
                return (DateTime)Properties.Settings.Default["LastDate"];
            }
            set {
                Properties.Settings.Default["LastDate"] = value;
                Properties.Settings.Default.Save();
            }
        }

        [Flags]
        public enum WarningFlags : byte {
            None = 0,
            Cpu = 1,
            Disk = 2,
        }

        public static string GetConfig(string valueName) {
            return ConfigurationManager.AppSettings[valueName];
        }

        /// <summary>
        /// クエリを作成します。
        /// </summary>
        /// <returns></returns>
        private static string getQuery() {
            if (GetConfig("ALLOWED_LOGTYPE") == null || GetConfig("ALLOWED_ENTRYLEVEL") == null) {
                throw new System.ArgumentException("ALLOWED_LOGTYPE又はALLOWED_ENTRYLEVELが空です");
            }
            List<string> ALLOWED_LOGTYPE = GetConfig("ALLOWED_LOGTYPE").Split(',').ToList();
            List<string> ALLOWED_ENTRYLEVEL = GetConfig("ALLOWED_ENTRYLEVEL").Split(',').ToList();

            int x = int.Parse(string.Format("{0:HH}", LastRunTime)) + 15;
            if (x > 23) x -= 24;
            string _hour = (x < 10) ? "0" + x : x + "";
            string _time = string.Format("{0:yyyy-MM-dd}T{1}:{0:mm}:{0:ss}", LastRunTime, _hour);
            string _level = string.Join(" or ", ALLOWED_ENTRYLEVEL.Select(_l => "Level=" + _l));
            string _query = "<QueryList><Query Id='0' Path='Application'>";
            foreach (string _type in ALLOWED_LOGTYPE) {
                _query += string.Format("<Select Path='{0}'>*[System[({1}) and TimeCreated[@SystemTime&gt;='{2}.000Z']]]</Select>", _type, _level, _time);
            }
            _query += "</Query></QueryList>";
            return _query;
        }

        /// <summary>
        /// Chatworkにメッセージを送信します。
        /// </summary>
        /// <param name="Token">chatworkのAPIキー</param>
        /// <param name="RoomId">chatworkのルームID</param>
        /// <param name="Title">タイトル</param>
        /// <param name="Message">本文</param>
        /// <returns></returns>
        public static async Task<HttpResponseMessage> SendMessage(string Token, string RoomId, string Title, string Message) {
            using (var client = new HttpClient()) {
                string PostStr = string.Format("body=[info][title]{0}[/title]{1}[/info]", Title, Message);
                client.DefaultRequestHeaders.Add("X-ChatWorkToken", Token);
                HttpContent content = new StringContent(PostStr);
                content.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");
                return await client.PostAsync(string.Format("https://api.chatwork.com/v1/rooms/{0}/messages", RoomId), content);
            }
        }

    }
    public static class EventLogReaderExtensions {
        public static IEnumerable<EventRecord> ReadAllEvents(this EventLogReader reader) {
            for (var record = reader.ReadEvent(); record != null; record = reader.ReadEvent()) {
                yield return record;
            }
        }
    }
}
