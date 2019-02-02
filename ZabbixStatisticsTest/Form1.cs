using Addin;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ZabbixStatisticsService;

namespace ZabbixStatisticsTest
{
    public partial class Form1 : Form
    {
        ServiceTest service;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ZabbixStatistics stat = new ZabbixStatistics();
            stat.НазваниеБазы = "MyTest";
            stat.Инициализация();
            stat.Начать("Test.TestProgram");
            Thread.Sleep(new Random().Next(200, 700));
            stat.Закончить();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            service = new ServiceTest();

            service.TestStart(new string[0]);

            button2.Enabled = false;
            button3.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            service.TestStop();
            service.Dispose();

            button2.Enabled = true;
            button3.Enabled = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ZabbixStatistics stat = new ZabbixStatistics();
            stat.НазваниеБазы = "MyTest";
            stat.Инициализация();
            stat.Начать("Test.TestProgram");
            Thread.Sleep(new Random().Next(200, 700));
            stat.Отменить();
            stat.Закончить();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ZabbixStatistics stat = new ZabbixStatistics();
            stat.НазваниеБазы = "MyTest";
            stat.Инициализация();
            stat.Начать("Test.TestProgram", -1000);
            Thread.Sleep(new Random().Next(200, 700));
            // Эмулируем некорректое завершение замера

            stat.Начать("Test.TestProgram", -1000);
            Thread.Sleep(new Random().Next(200, 700));
            stat.Закончить();
        }
    }

    class ServiceTest : Service
    {
        public void TestStart(string[] args)
        {
            OnStart(args);
        }

        public void TestStop()
        {
            OnStop();
        }

    }
}
