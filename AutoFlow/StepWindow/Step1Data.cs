using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoFlow.StepWindow
{
    class Step1Data
    {
        public delegate void SendValueEventHandler(string e);
        static public event SendValueEventHandler SendValueEventHandler1;
        static public event SendValueEventHandler SendValueEventHandler2;

        static private string data1;
        static private string data2;

        static public string Step1_data1
        {
            get
            {
                return data1;
            }
            set
            {
                data1 = value;
                SendValueEventHandler1(data1);
            }
        }
        static public string Step1_data2
        {
            get
            {
                return data2;
            }
            set
            {
                data2 = value;
                SendValueEventHandler2(data2);
            }
        }
        public delegate void CheckSendValueEventHandler(bool e);
        static public event CheckSendValueEventHandler CheckSendValueEventHandler1;
        static private bool bool1;
        static public bool Step1_bool1
        {
            get
            {
                return bool1;
            }
            set
            {
                bool1 = value;
                CheckSendValueEventHandler1(bool1);
            }
        }

    }
}
