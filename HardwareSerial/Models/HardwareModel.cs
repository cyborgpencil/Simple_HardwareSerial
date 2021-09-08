using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HardwareSerial.Models
{
    public class HardwareModel
    {
        public string Hostname { get; set; } = "NA";
        public string BoardSerialName { get; set; } = "NA";
        public List<MonitorInfo> Monitors { get; set; } = new List<MonitorInfo>();
    }
}
