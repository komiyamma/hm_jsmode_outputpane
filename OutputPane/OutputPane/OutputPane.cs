using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using HmNetCOM;
using static HmNetCOM.Hm;

namespace OutputPane
{

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("5BC1D050-BF30-42C5-B2A5-DB6918BF9813")]

    public class OutputPane : IOutputPane
    {
        public int OutputPane_SendMessage(Object command_id)
        {
            try
            {
                int id = (int)(dynamic)(command_id);
                IntPtr ret = Hm.OutputPane.SendMessage(id);
                return ret.ToInt32();
            } catch(Exception e)
            {
                Hm.OutputPane.Output(e.Message + "\r\n");
            }
            return 0;
        }
    }

    public interface IOutputPane
    {
        int OutputPane_SendMessage(Object command_id);
    }
}
