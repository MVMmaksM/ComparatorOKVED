using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Comparator.OKVED.Model;

namespace Comparator.OKVED.Comparator
{
    class CompareOKVED
    {
        public IEnumerable<ModelResultHozChist> CompareChistHozOkved(IEnumerable<ModelPBD> collectionPBD)
        {
            var delokvedAG = DelOkvedAG(collectionPBD).Where(a => a.OtchMes != null || a.OtchKvart != null).GroupBy(a => new { a.OKPO, a.OKVEDHoz, a.OKVEDChist }).Select(a => a.First());
            var hozEqualsChist = GetRowsOkvedHozEqualsChist(delokvedAG);

            var delHozSubStringChist = delokvedAG.Where(a => !a.OKVEDChist.Contains(a.OKVEDHoz));
            var delChistSubStringHoz = delHozSubStringChist.Where(a => !a.OKVEDHoz.Contains(a.OKVEDChist)).ToList();

            foreach (var row in hozEqualsChist)
            {
                delChistSubStringHoz.RemoveAll(a => a.OKPO == row.OKPO);
            }

            return delChistSubStringHoz.Select(a => new ModelResultHozChist()
            {
                Period = a.Period,
                OKPO = a.OKPO,
                Name = a.Name,
                OKATO = a.OKATO,
                OKVEDHoz = a.OKVEDHoz,
                OKVEDChist = a.OKVEDChist
            });
        }
        public IEnumerable<ModelResultChist> CompareChistOkved(IEnumerable<ModelPBD> dataCurPerPBD, IEnumerable<ModelPBD> dataPrevPerPBD)
        {
            var delAGOkvedCurPer = DelOkvedAG(dataCurPerPBD);
            var delAGOkvedPrevPer = DelOkvedAG(dataPrevPerPBD);

            var groupOkpoOkvedCurPer = delAGOkvedCurPer.GroupBy(a => new { a.OKPO, a.OKVEDChist }).Select(a => a
              .First());

            var groupOkpoOkvedPrevPer = delAGOkvedPrevPer.GroupBy(a => new { a.OKPO, a.OKVEDChist }).Select(a => a
              .First());

            var leftJoinCurPer = from cur in groupOkpoOkvedCurPer   //left join для текущего периода
                                 join prev in groupOkpoOkvedPrevPer
                                 on new { OKPO = cur.OKPO, OKVED = cur.OKVEDChist } equals new { OKPO = prev.OKPO, OKVED = prev.OKVEDChist }
                                 into data_A
                                 from data_B in data_A.DefaultIfEmpty(new ModelPBD())
                                 select new ModelResultChist()
                                 {
                                     Period = cur.Period,
                                     OKPO = cur.OKPO,
                                     Name = cur.Name,
                                     OKATO = cur.OKATO,
                                     OKVEDCurPer = cur.OKVEDChist,
                                     OKVEDPrevPer = data_B.OKVEDChist
                                 };

            var notInPrevPer = leftJoinCurPer.Where(a => string.IsNullOrEmpty(a.OKVEDPrevPer));

            var leftJoinPrevPer = from prev in groupOkpoOkvedPrevPer  //left join для предыдущего периода
                                  join cur in groupOkpoOkvedCurPer
                                  on new { OKPO = prev.OKPO, OKVED = prev.OKVEDChist } equals new { OKPO = cur.OKPO, OKVED = cur.OKVEDChist }
                                  into data_A
                                  from data_B in data_A.DefaultIfEmpty(new ModelPBD())
                                  select new ModelResultChist()
                                  {
                                      Period = prev.Period,
                                      OKPO = prev.OKPO,
                                      Name = prev.Name,
                                      OKATO = prev.OKATO,
                                      OKVEDCurPer = data_B.OKVEDChist,
                                      OKVEDPrevPer = prev.OKVEDChist
                                  };

            var notInCurPer = leftJoinPrevPer.Where(a => string.IsNullOrEmpty(a.OKVEDCurPer));
                      
            var unionCurPrevPer = notInPrevPer.Union(notInCurPer);

            return unionCurPrevPer;
        }
        private IEnumerable<ModelPBD> GetRowsOkvedHozEqualsChist(IEnumerable<ModelPBD> collectionPBD) => collectionPBD.Where(a => a.OKVEDHoz == a.OKVEDChist);
        private IEnumerable<ModelPBD> DelOkvedAG(IEnumerable<ModelPBD> collectionPBD) => collectionPBD.Where(a => a.OKVEDChist != "101.АГ");
        private PropertyInfo GetPropertyInfoRadioButton(string radioButtonOrderBy)
        {
            string paramName = string.Empty;

            var modelResultChist = new ModelResultChist();
            switch (radioButtonOrderBy)
            {
                case "ОКПО":
                    paramName = nameof(modelResultChist.OKPO);
                    break;
                case "Наименованию предприятия":
                    paramName = nameof(modelResultChist.Name);
                    break;
                case "ОКАТО":
                    paramName = nameof(modelResultChist.OKATO);
                    break;
                default:
                    break;
            }

            return typeof(ModelResultChist).GetProperty(paramName);
        }
    }
}