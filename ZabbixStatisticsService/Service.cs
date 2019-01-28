using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using IniFiles;
using System.Threading;
using System.IO;
using System.Reflection;
using System.Data.SqlClient;
using Ysq.Zabbix;
using NLog;

namespace ZabbixStatisticsService
{
    public partial class Service : ServiceBase
    {
        ServiceThread thr;
        private string workDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        private Logger logger = LogManager.GetCurrentClassLogger();

        public Service()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            #region Чтение настроек
            bool iniError = false;
            IniFile INI = new IniFile(workDirectory + "\\config.ini");

            if (!INI.KeyExists("SqlServer", "SQL")) { INI.Write("SQL", "SqlServer", "mysqlserver"); iniError = true; }
            if (!INI.KeyExists("SqlBaseName", "SQL")) { INI.Write("SQL", "SqlBaseName", "sqlBaseName"); iniError = true; }
            if (!INI.KeyExists("SqlUser", "SQL")) { INI.Write("SQL", "SqlUser", "myUserName"); iniError = true; }
            if (!INI.KeyExists("SqlPass", "SQL")) { INI.Write("SQL", "SqlPass", "myUserPassword"); iniError = true; }

            if (!INI.KeyExists("ServiceThreadSleep", "OTHER")) { INI.Write("OTHER", "ServiceThreadSleep", "60"); iniError = true; }

            if (!INI.KeyExists("ZabbixServer", "ZABBIX")) { INI.Write("ZABBIX", "ZabbixServer", "zabbix.mynet.com"); iniError = true; }
            if (!INI.KeyExists("ZabbixPort", "ZABBIX")) { INI.Write("ZABBIX", "ZabbixPort", "10051"); iniError = true; }
            if (!INI.KeyExists("ZabbixNodeName", "ZABBIX")) { INI.Write("ZABBIX", "ZabbixNodeName", "1c"); iniError = true; }



            if (iniError)
            {
                logger.Error("Ошибка чтения настроек!");
                throw new Exception("Ошибка чтения настроек!");
            }
            else
            {
                try
                {
                    thr = new ServiceThread(logger);
                    thr.sqlServer = INI.ReadINI("SQL", "SqlServer");
                    thr.sqlBaseName = INI.ReadINI("SQL", "SqlBaseName");
                    thr.sqlUser = INI.ReadINI("SQL", "SqlUser");
                    thr.sqlPass = INI.ReadINI("SQL", "SqlPass");

                    thr.serviceThreadSleep = Convert.ToInt32(INI.ReadINI("OTHER", "ServiceThreadSleep"));

                    thr.zabbixServer = INI.ReadINI("ZABBIX", "ZabbixServer");
                    thr.zabbixPort = Convert.ToInt32(INI.ReadINI("ZABBIX", "ZabbixPort"));
                    thr.zabbixNodeName = INI.ReadINI("ZABBIX", "ZabbixNodeName");

                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Ошибка чтения настроек!");
                    throw new Exception("Ошибка чтения настроек! " + ex.Message);
                }
            }
            #endregion

            thr.Start();
            logger.Info("Служба успешно запущена");
        }

        protected override void OnStop()
        {
            logger.Info("Остановка службы");
            thr.Stop();
            logger.Info("Служба успешно остановлена");
        }
    }

    public class ServiceThread
    {
        public string sqlServer = "mysqlserver";
        public string sqlBaseName = "sqlBaseName";
        public string sqlUser = "sqlUser";
        public string sqlPass = "sqlPass";
        public int serviceThreadSleep = 60;
        public string zabbixServer = "zabbix.mynet.com";
        public int zabbixPort = 10051;
        public string zabbixNodeName = "1c";
        Logger logger;
        Thread thr;
        bool serviceStop = false;

        public ServiceThread(Logger logger)
        {
            this.logger = logger;
        }
        public void Start()
        {
            thr = new Thread(ZabbixSender);
            thr.IsBackground = true;
            thr.Start();
        }
        public void Stop()
        {
            serviceStop = true;
            thr.Join();
        }

        private void ZabbixSender()
        {

            string connectionString = new SqlConnectionStringBuilder()
            {
                DataSource = sqlServer,
                InitialCatalog = sqlBaseName,
                UserID = sqlUser,
                Password = sqlPass
            }.ConnectionString;

            Sender sender = new Sender(zabbixServer, zabbixPort);

            while (!serviceStop)
            {
                try
                {
                    using (SqlConnection oConnection = new SqlConnection(connectionString))
                    {
                        oConnection.Open();
                        SqlCommand command = new SqlCommand(
                            @"SELECT [Basename],[Param], AVG([Value]) as 'Value' FROM[Statistics]
                          GROUP BY[Basename],[Param]
                          TRUNCATE TABLE[Statistics]
                          ", oConnection);
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                string send = String.Format("1C.Statistics[{0},{1}]", reader["BaseName"].ToString().Trim(), reader["Param"].ToString().Trim());
                                try
                                {
                                    SenderResponse response = sender.Send(zabbixNodeName, send, reader["Value"].ToString());
                                    logger.Debug(send);
                                    logger.Debug(response.Response);
                                    logger.Debug(response.Info);
                                }
                                catch (Exception ex)
                                {
                                    logger.Error(ex, "Ошибка отправки статистики в zabbix");
                                }
                            }
                        }
                        oConnection.Close();
                    }
                }
                catch(Exception ex) {
                    logger.Error(ex,"Ошибка запроса данных из SQL");
                }
                #region Задержка
                for (int i = 1; i < serviceThreadSleep; i++)
                {
                    if (serviceStop) return;
                    Thread.Sleep(1000);
                }
                #endregion
            }
        }
    }
}
