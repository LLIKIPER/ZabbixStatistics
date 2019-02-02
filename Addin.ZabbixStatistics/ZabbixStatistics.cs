using IniFiles;
using Microsoft.Win32;
using NLog;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace Addin
{
    [Guid("CF446236-2E44-4805-805C-068C8E81A953")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IZabbixStatistics
    {
        void Инициализация();
        string НазваниеБазы { get; set; }
        void Начать(string key, int failed_value=-1000);
        void Отменить();
        void Закончить();
    }


    [ProgId("Addin.ZabbixStatistics")]
    [ClassInterface(ClassInterfaceType.AutoDual), ComSourceInterfaces(typeof(IZabbixStatistics))]
    [Guid("4F273AA3-A95A-4CD4-95FA-3FFBE1B8CAEB")]
    [ComVisible(true)]
    public class ZabbixStatistics : IZabbixStatistics
    {
        private string sqlServer = "mysqlserver";
        private string sqlBaseName = "sqlBaseName";
        private string sqlUser = "sqlUser";
        private string sqlPass = "sqlPass";
        private string baseName = "MyBase";
        private int threadSleep = 10 * 1000;
        private int _batchSize = 100;

        private string workDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        private ZabbixList list = new ZabbixList();
        private Thread thr = null;
        private string _key="";
        private int _failed_value = -1000;
        private DateTime _startTime = DateTime.MinValue;

        private Logger logger = LogManager.GetCurrentClassLogger();

        public ZabbixStatistics() { }


        [ComVisible(true)]
        public void Инициализация()
        {
            #region Чтение настроек
            IniFile INI = new IniFile(workDirectory + "\\config.ini");

            bool iniError = false;

            if (!INI.KeyExists("SqlServer", "SQL")) { INI.Write("SQL", "SqlServer", "mysqlserver"); iniError = true; }
            if (!INI.KeyExists("SqlBaseName", "SQL")) { INI.Write("SQL", "SqlBaseName", "sqlBaseName"); iniError = true; }
            if (!INI.KeyExists("SqlUser", "SQL")) { INI.Write("SQL", "SqlUser", "myUserName"); iniError = true; }
            if (!INI.KeyExists("SqlPass", "SQL")) { INI.Write("SQL", "SqlPass", "myUserPassword"); iniError = true; }
            if (!INI.KeyExists("ThreadSleep", "OTHER")) { INI.Write("OTHER", "ThreadSleep", "10"); iniError = true; }


            if (iniError)
            {
                logger.Error("Ошибка чтения настроек!");
                throw new Exception("Ошибка чтения настроек!");
            }
            else
            {
                try
                {
                    sqlServer = INI.ReadINI("SQL", "SqlServer");
                    sqlBaseName = INI.ReadINI("SQL", "SqlBaseName");
                    sqlUser = INI.ReadINI("SQL", "SqlUser");
                    sqlPass = INI.ReadINI("SQL", "SqlPass");
                    threadSleep = 1000 * Convert.ToInt32(INI.ReadINI("OTHER", "ThreadSleep"));
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Ошибка чтения настроек!");
                    throw new Exception("Ошибка чтения настроек! " + ex.Message);
                }
            }
            #endregion

            thr = new Thread(SendStatistics);
            thr.IsBackground = true;
            thr.Start();
        }

        [ComVisible(true)]
        public string НазваниеБазы { get { return baseName; } set { baseName = value; } }


        [ComVisible(true)]
        public void Начать(string key, int failed_value = -1000)
        {
            if (thr == null)
                throw new Exception("Перед методом \"Начать\", нужно выполнить метод \"Инициализация\"!");

            if (key.Trim() == "")
                throw new Exception("Обязательно нужно передать не пустой ключ в метод \"Начать\"");

            if (_key.Trim() != "" && _startTime != DateTime.MinValue)
            {
                //Предыдущий замер не был завершен корректно, отправляем -1 как показатель ошибки/отмены
                list.add(_key, _failed_value.ToString());
            }
            // Начало нового замера
            _key = key;
            _startTime = DateTime.Now;
            _failed_value = failed_value;
        }

        [ComVisible(true)]
        public void Закончить()
        {
            if (_startTime == DateTime.MinValue) return;
            if (_key.Trim() == "") return;
            TimeSpan ts = DateTime.Now - _startTime;
            list.add(_key, Math.Truncate(ts.TotalMilliseconds).ToString());

            Отменить();
        }

        [ComVisible(true)]
        public void Отменить()
        {
            _startTime = DateTime.MinValue;
            _key = "";
        }

        private void SendStatistics()
        {
            string connectionString = new SqlConnectionStringBuilder()
            {
                DataSource = sqlServer,
                InitialCatalog = sqlBaseName,
                UserID = sqlUser,
                Password = sqlPass
            }.ConnectionString;

            while (true)
            {
                if (list.count() > 0)
                {
                    try
                    {
                        using (SqlConnection oConnection = new SqlConnection(connectionString))
                        {
                            oConnection.Open();
                            using (SqlTransaction oTransaction = oConnection.BeginTransaction())
                            {
                                using (SqlCommand oCommand = oConnection.CreateCommand())
                                {
                                    oCommand.Transaction = oTransaction;
                                    oCommand.CommandType = CommandType.Text;
                                    oCommand.CommandText = "INSERT INTO [Statistics] ([Time], [Basename], [Param], [Value]) VALUES (@Time, @Basename, @Param, @Value);";
                                    oCommand.Parameters.Add(new SqlParameter("@Time", SqlDbType.DateTime));
                                    oCommand.Parameters.Add(new SqlParameter("@Basename", SqlDbType.NChar));
                                    oCommand.Parameters.Add(new SqlParameter("@Param", SqlDbType.NChar));
                                    oCommand.Parameters.Add(new SqlParameter("@value", SqlDbType.Int));
                                    try
                                    {
                                        List<ZabbixClass> sendList = list.getListAndClear();
                                        foreach (ZabbixClass cur in sendList)
                                        {
                                            oCommand.Parameters[0].Value = cur.Time;
                                            oCommand.Parameters[1].Value = baseName;
                                            oCommand.Parameters[2].Value = cur.Param;
                                            oCommand.Parameters[3].Value = cur.Value;
                                            if (oCommand.ExecuteNonQuery() != 1)
                                            {
                                                //'handled as needed, 
                                                //' but this snippet will throw an exception to force a rollback
                                                throw new InvalidProgramException();
                                            }

                                        }
                                        oTransaction.Commit();
                                    }
                                    catch (Exception)
                                    {
                                        oTransaction.Rollback();
                                        throw;
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, "Ошибка отправки замеров в SQL");
                    }
                }

                Thread.Sleep(threadSleep);
            }
        }


        #region Функции регистрации компоненты
        [ComRegisterFunction()]
        public static void RegisterClass(string key)
        {
            // Strip off HKEY_CLASSES_ROOT\ from the passed key as I don't need it
            StringBuilder sb = new StringBuilder(key);
            sb.Replace(@"HKEY_CLASSES_ROOT\", "");
            // Open the CLSID\{guid} key for write access
            RegistryKey k = Registry.ClassesRoot.OpenSubKey(sb.ToString(), true);
            // And create	the	'Control' key -	this allows	it to show up in
            // the ActiveX control container
            RegistryKey ctrl = k.CreateSubKey("Control");
            ctrl.Close();
            // Next create the CodeBase entry	- needed if	not	string named and GACced.
            RegistryKey inprocServer32 = k.OpenSubKey("InprocServer32", true);
            inprocServer32.SetValue("CodeBase", Assembly.GetExecutingAssembly().CodeBase);
            inprocServer32.Close();
            // Finally close the main	key
            k.Close();
            Console.WriteLine(@"
**************************************************************
            Успешно зарегистрирована!
**************************************************************
            ");
        }
        [ComUnregisterFunction()]
        public static void UnregisterClass(string key)
        {
            StringBuilder sb = new StringBuilder(key);
            sb.Replace(@"HKEY_CLASSES_ROOT\", "");
            // Open	HKCR\CLSID\{guid} for write	access
            RegistryKey k = Registry.ClassesRoot.OpenSubKey(sb.ToString(), true);
            // Delete the 'Control'	key, but don't throw an	exception if it	does not exist
            k.DeleteSubKey("Control", false);
            // Next	open up	InprocServer32
            //RegistryKey	inprocServer32 = 
            k.OpenSubKey("InprocServer32", true);
            // And delete the CodeBase key,	again not throwing if missing
            k.DeleteSubKey("CodeBase", false);
            // Finally close the main key
            k.Close();
            Console.WriteLine(@"
**************************************************************
            Регистрация успешно отменена!
**************************************************************
            ");
        }
        #endregion
    }


    #region Дополнительные классы
    public class ZabbixList
    {
        object _lock = new object();
        List<ZabbixClass> _list;

        public ZabbixList()
        {
            _list = new List<ZabbixClass>();
        }

        public List<ZabbixClass> getListAndClear()
        {
            lock (_lock)
            {
                List<ZabbixClass> _resultList = new List<ZabbixClass>(_list);
                _list.Clear();
                return _resultList;
            }
        }
        public void add(string Key, string Value)
        {
            lock (_lock)
            {
                _list.Add(new ZabbixClass(Key, Value));
            }
        }

        public int count()
        {
            lock (_lock)
            {
                return _list.Count();
            }
        }

    }

    public class ZabbixClass : IEquatable<ZabbixClass>
    {
        DateTime _Time;
        string _Param;
        string _Value;

        public ZabbixClass(string Param, string Value)
        {
            _Time = DateTime.Now;
            _Param = Param;
            _Value = Value;
        }
        public DateTime Time
        {
            get { return _Time; }
            set { _Time = value; }
        }
        public string Param
        {
            get { return _Param; }
            set { _Param = value; }
        }

        public string Value
        {
            get { return _Value; }
            set { _Value = value; }
        }

        #region Override функции
        public override int GetHashCode() { return 0; }
        public override string ToString()
        {
            return "Время: " + _Time.ToString("dd.mm.yyy HH:MM:ss") + " Параметр: " + _Param + " значение: " + _Value;
        }
        public bool Equals(ZabbixClass other)
        {
            if (other == null) return false;
            if (
                (this._Param.Equals(other._Param)) &&
                (this._Value.Equals(other._Value))
                )
                return true;
            else return false;
        }
        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            ZabbixClass objAsGetЗапрос = obj as ZabbixClass;
            if (objAsGetЗапрос == null) return false;
            else return Equals(objAsGetЗапрос);
        }
        #endregion

    }
    #endregion

}
