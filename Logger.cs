using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Eplan.EplApi.Base;

namespace CENTEC.EplAPI.Service
{
    internal static class Logger
    {
        public static void SendMSGToEplanLog(String msg)
        {
            BaseException exc = new BaseException(msg, MessageLevel.Message);
            exc.FixMessage();
        }

        //public interface ILogBoxable
        //{
        //    void ToLog(String msg);
        //}
    }
}
