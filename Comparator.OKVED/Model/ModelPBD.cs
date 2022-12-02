using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Comparator.OKVED.Model
{
    class ModelPBD
    {
        public string Period { get; set; }
        public string OKATO { get; set; }
        public string OKPO { get; set; }
        public string Name { get; set; }
        public string KodPokaz { get; set; }
        public string OKVEDHoz { get; set; }
        public string OKVEDChist { get; set; }
        public double? OtchMes { get; set; }
        public double? PredMes { get; set; }
        public double? SovMesPrGod { get; set; }
        public double? OtchKvart { get; set; }
        public double? PredKvart { get; set; }
        public double? SovKvartPrGod { get; set; }
    }
}
