/*
上次研究了.Net版本的OPC API dll，这次我采用OPCDAAuto.dll来介绍使用方法。

以下为我的源代码，有详细的注释无需我多言。
编译平台：VS2008SP1、WINXP、KEPServer
除此之外，我也安装了西门子的Net2006和Step7，其中Net2006是负责OPC的，可能会在系统中创建一些dll之类的，并提供几个OPC服务器
以下是我Program.cs（基于控制台的）的全部内容，仍旧采用C#语言：
*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Collections;
using OPCAutomation;
using System.Threading;

namespace OPCDAAutoTest
{
    class Tester
    {
        static void Main(string[] args)
        {
            Tester tester=new Tester();
            tester.work();
        }
        #region 私有变量
        /// <summary>
        /// OPCServer Object
        /// </summary>
        OPCServer MyServer;
        /// <summary>
        /// OPCGroups Object
        /// </summary>
        OPCGroups MyGroups;
        /// <summary>
        /// OPCGroup Object
        /// </summary>
        OPCGroup MyGroup;
        OPCGroup MyGroup2;
        /// <summary>
        /// OPCItems Object
        /// </summary>
        OPCItems MyItems;
        OPCItems MyItems2;
        /// <summary>
        /// OPCItem Object
        /// </summary>
        OPCItem[] MyItem;
        OPCItem[] MyItem2;
        /// <summary>
        /// 主机IP
        /// </summary>
        string strHostIP = "";
        /// <summary>
        /// 主机名称
        /// </summary>
        string strHostName = "";
        /// <summary>
        /// 连接状态
        /// </summary>
        bool opc_connected = false;
        /// <summary>
        /// 客户端句柄
        /// </summary>
        int itmHandleClient = 0;
        /// <summary>
        /// 服务端句柄
        /// </summary>
        int itmHandleServer = 0;
        #endregion
        //测试用工作方法
        public void work()
        {
            //初始化item数组
            MyItem = new OPCItem[4];
            MyItem2 = new OPCItem[4];

            GetLocalServer();
            //ConnectRemoteServer("TX1", "KEPware.KEPServerEx.V4");//用计算机名的局域网
            //ConnectRemoteServer("192.168.1.35", "KEPware.KEPServerEx.V4");//用IP的局域网
            if (ConnectRemoteServer("", "KEPware.KEPServerEx.V4"))//本机
            {
                if (CreateGroup())
                {
                    Thread.Sleep(500);//暂停线程以让DataChange反映，否则下面的同步读可能读不到
                    //以下同步写
                    MyItem[3].Write("I love you!");//同步写
                    MyItem[2].Write(true);//同步写
                    MyItem[1].Write(-100);//同步写
                    MyItem[0].Write(120);//同步写              
   
                    //以下同步读
                    object ItemValues;  object Qualities; object TimeStamps;//同步读的临时变量：值、质量、时间戳
                    MyItem[0].Read(1,out ItemValues,out Qualities,out TimeStamps);//同步读，第一个参数只能为1或2
                    int q0 = Convert.ToInt32(ItemValues);//转换后获取item值
                    MyItem[1].Read(1, out ItemValues, out Qualities, out TimeStamps);//同步读，第一个参数只能为1或2
                    int q1 = Convert.ToInt32(ItemValues);//转换后获取item值
                    MyItem[2].Read(1, out ItemValues, out Qualities, out TimeStamps);//同步读，第一个参数只能为1或2
                    bool q2 = Convert.ToBoolean(ItemValues);//转换后获取item值
                    MyItem[3].Read(1, out ItemValues, out Qualities, out TimeStamps);//同步读，第一个参数只能为1或2
                    string q3 = Convert.ToString(ItemValues);//转换后获取item值，为防止读到的值为空，不用ItemValues.ToString()

                    Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                    Console.WriteLine("0-{0},1-{1},2-{2},3-{3}",q0,q1,q2,q3);

                    //以下为异步写
                    //异步写时，Array数组从下标1开始而非0！
                    int[] temp = new int[] { 0,MyItem[0].ServerHandle,MyItem[1].ServerHandle,MyItem[2].ServerHandle, MyItem[3].ServerHandle };
                    Array serverHandles = (Array)temp;
                    object[] valueTemp = new object[5] { "",255,520,true, "Love" };
                    Array values = (Array)valueTemp;
                    Array Errors;
                    int cancelID;
                    MyGroup.AsyncWrite(4, ref serverHandles, ref values, out Errors, 1, out cancelID);//第一参数为item数量
                    //由于MyGroup2没有订阅，所以以下这句运行时将会出错！
                    //MyGroup2.AsyncWrite(4, ref serverHandles, ref values, out Errors, 1, out cancelID);

                    //以下异步读
                    MyGroup.AsyncRead(4, ref serverHandles, out Errors, 1, out cancelID);//第一参数为item数量


                    /*MyItem[0] = MyItems.AddItem("BPJ.Db1.dbb96", 0);//byte
                    MyItem[1] = MyItems.AddItem("BPJ.Db1.dbw10", 1);//short
                    MyItem[2] = MyItems.AddItem("BPJ.Db16.dbx0", 2);//bool
                    MyItem[3] = MyItems.AddItem("BPJ.Db11.S0", 3);//string*/

 

 

                    Console.WriteLine("************************************** hit <return> to Disconnect...");
                    Console.ReadLine();  
                    //释放所有组资源
                    MyServer.OPCGroups.RemoveAll();
                    //断开服务器
                    MyServer.Disconnect();
                }
            }


            //END
            Console.WriteLine("************************************** hit <return> to close...");
            Console.ReadLine();  
        }

  //枚举本地OPC服务器
        private void GetLocalServer()
        {
            //获取本地计算机IP,计算机名称
            strHostName = Dns.GetHostName();
            //或者通过局域网内计算机名称
            strHostName = "TX1";

            //获取本地计算机上的OPCServerName
            try
            {
                MyServer = new OPCServer();
                object serverList = MyServer.GetOPCServers(strHostName);

                foreach (string server in (Array)serverList)
                {
                    //cmbServerName.Items.Add(turn);
                    Console.WriteLine("本地OPC服务器：{0}", server);
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("枚举本地OPC服务器出错：{0}",err.Message);
            }
        }
        //连接OPC服务器
        /// <param name="remoteServerIP">OPCServerIP</param>
        /// <param name="remoteServerName">OPCServer名称</param>
        private bool ConnectRemoteServer(string remoteServerIP, string remoteServerName)
        {
            try
            {
                MyServer.Connect(remoteServerName, remoteServerIP);//连接本地服务器：服务器名+主机名或IP

                if (MyServer.ServerState == (int)OPCServerState.OPCRunning)
                {
                    Console.WriteLine("已连接到：{0}",MyServer.ServerName);
                }
                else
                {
                    //这里你可以根据返回的状态来自定义显示信息，请查看自动化接口API文档
                    Console.WriteLine("状态：{0}",MyServer.ServerState.ToString());
                }
                MyServer.ServerShutDown+=ServerShutDown;//服务器断开事件
            }
            catch (Exception err)
            {
                Console.WriteLine("连接远程服务器出现错误：{0}" + err.Message);
                return false;
            }
            return true;
        }
        //创建组
        private bool CreateGroup()
        {
            try
            {
                MyGroups = MyServer.OPCGroups;
                MyGroup = MyServer.OPCGroups.Add("测试");//添加组
                MyGroup2 = MyGroups.Add("测试2");
                OPCGroup MyGroup3 = MyGroups.Add("测试3");//测试删除组
                //以下设置组属性
                {
                    MyServer.OPCGroups.DefaultGroupIsActive = true;//激活组。
                    MyServer.OPCGroups.DefaultGroupDeadband = 0;// 死区值，设为0时，服务器端该组内任何数据变化都通知组。
                    MyServer.OPCGroups.DefaultGroupUpdateRate = 200;//默认组群的刷新频率为200ms
                    MyGroup.UpdateRate = 100;//刷新频率为1秒。
                    MyGroup.IsSubscribed = true;//使用订阅功能，即可以异步，默认false
                }       

                MyGroup.DataChange += new DIOPCGroupEvent_DataChangeEventHandler(GroupDataChange);
                MyGroup.AsyncWriteComplete += new DIOPCGroupEvent_AsyncWriteCompleteEventHandler(GroupAsyncWriteComplete);
                MyGroup.AsyncReadComplete += new DIOPCGroupEvent_AsyncReadCompleteEventHandler(GroupAsyncReadComplete);
                //由于MyGroup2.IsSubscribed是false，即没有订阅，所以以下的DataChange回调事件不会发生！
                MyGroup2.DataChange += new DIOPCGroupEvent_DataChangeEventHandler(GroupDataChange2);
                MyGroup.AsyncWriteComplete += new DIOPCGroupEvent_AsyncWriteCompleteEventHandler(GroupAsyncWriteComplete); 
                 
                MyServer.OPCGroups.Remove("测试3");//移除组
                AddGroupItems();//设置组内items         
            }
            catch (Exception err)
            {
                Console.WriteLine("创建组出现错误：{0}", err.Message);
                return false;
            }
            return true;
        }
        private void AddGroupItems()//添加组
        {
            //itmHandleServer;
            MyItems = MyGroup.OPCItems;
            MyItems2 = MyGroup2.OPCItems;

            //添加item
            MyItem[0] = MyItems.AddItem("BPJ.Db1.dbb96", 0);//byte
            MyItem[1] = MyItems.AddItem("BPJ.Db1.dbw10", 1);//short
            MyItem[2] = MyItems.AddItem("BPJ.Db16.dbx0", 2);//bool
            MyItem[3] = MyItems.AddItem("BPJ.Db11.S0", 3);//string
            //移除组内item
            Array Errors;      
            int []temp=new int[]{0,MyItem[3].ServerHandle};
            Array serverHandle = (Array)temp;
            MyItems.Remove(1, ref serverHandle, out Errors);
            MyItem[3] = MyItems.AddItem("BPJ.Db11.S0", 3);//string

            MyItem2[0] = MyItems2.AddItem("BPJ.Db1.dbb96", 0);//byte
            MyItem2[1] = MyItems2.AddItem("BPJ.Db1.dbw10", 1);//short
            MyItem2[2] = MyItems2.AddItem("BPJ.Db16.dbx0", 2);//bool
            MyItem2[3] = MyItems2.AddItem("BPJ.Db11.S0", 3);//string

        }
        public void ServerShutDown(string Reason)//服务器先行断开
        {
            Console.WriteLine("服务器已经先行断开！");
        }
        /// <summary>
        /// 每当项数据有变化时执行的事件
        /// </summary>
        /// <param name="TransactionID">处理ID</param>
        /// <param name="NumItems">项个数</param>
        /// <param name="ClientHandles">项客户端句柄</param>
        /// <param name="ItemValues">TAG值</param>
        /// <param name="Qualities">品质</param>
        /// <param name="TimeStamps">时间戳</param>1    `
        void GroupDataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {
            Console.WriteLine("++++++++++++++++DataChanged+++++++++++++++++++++++");
            /*for (int i = 1; i <= NumItems; i++)
            {
                Console.WriteLine("item值：{0}", ItemValues.GetValue(i).ToString());
                //Console.WriteLine("item句柄：{0}", ClientHandles.GetValue(i).ToString());
                //Console.WriteLine("item质量：{0}", Qualities.GetValue(i).ToString());
                //Console.WriteLine("item时间戳：{0}", TimeStamps.GetValue(i).ToString());
                //Console.WriteLine("item类型：{0}", ItemValues.GetValue(i).GetType().FullName);
            }*/
        }
        void GroupDataChange2(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {
            Console.WriteLine("----------------------DataChanged2------------------");
            /*for (int i = 1; i <= NumItems; i++)
            {
                Console.WriteLine("item2值：{0}", ItemValues.GetValue(i).ToString());
                //Console.WriteLine("item2质量：{0}", Qualities.GetValue(i).ToString());
                //Console.WriteLine("item2时间戳：{0}", TimeStamps.GetValue(i).ToString());             
            }*/
        }
       

 /// <summary>
        /// 异步写完成
        /// 运行时，Array数组从下标1开始而非0！
        /// </summary>
        /// <param name="TransactionID"></param>
        /// <param name="NumItems"></param>
        /// <param name="ClientHandles"></param>
        /// <param name="Errors"></param>
        void GroupAsyncWriteComplete(int TransactionID, int NumItems, ref Array ClientHandles, ref Array Errors)
        {
            Console.WriteLine("%%%%%%%%%%%%%%%%AsyncWriteComplete%%%%%%%%%%%%%%%%%%%");
            /*for (int i = 1; i <= NumItems; i++)
            {
                Console.WriteLine("Tran：{0}   ClientHandles：{1}   Error：{2}", TransactionID.ToString(), ClientHandles.GetValue(i).ToString(), Errors.GetValue(i).ToString());
            }*/
        }

        /// <summary>
        /// 异步读完成
        /// 运行时，Array数组从下标1开始而非0！
        /// </summary>
        /// <param name="TransactionID"></param>
        /// <param name="NumItems"></param>
        /// <param name="ClientHandles"></param>
        /// <param name="ItemValues"></param>
        /// <param name="Qualities"></param>
        /// <param name="TimeStamps"></param>
        /// <param name="Errors"></param>
        void GroupAsyncReadComplete(int TransactionID, int NumItems, ref System.Array ClientHandles, ref System.Array ItemValues, ref System.Array Qualities, ref System.Array TimeStamps, ref System.Array Errors)
        {
            Console.WriteLine("****************GroupAsyncReadComplete*******************");
            for (int i = 1; i <= NumItems; i++)
            {
                //Console.WriteLine("Tran：{0}   ClientHandles：{1}   Error：{2}", TransactionID.ToString(), ClientHandles.GetValue(i).ToString(), Errors.GetValue(i).ToString());
                Console.WriteLine("Vaule：{0}",Convert.ToString(ItemValues.GetValue(i)));
            }
        }
    }
}
