using ComparatorOKVED.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ComparatorOKVED.Comaprator
{
    class ComparatorOKVED
    {
        public static List<ModelResult> CompareOKVED(List<Model> listPred, List<Model> listTek)
        {
            try
            {
                var resultPred = listPred.Where(z => z.OKVED != "101.АГ").GroupBy(z => new { z.OKPO, z.OKVED }).Select(z => z.First()).ToList();
                var resultTek = listTek.Where(z => z.OKVED != "101.АГ").GroupBy(z => new { z.OKPO, z.OKVED }).Select(z => z.First()).ToList();

                var leftJoinPred = from p in resultPred   //left join для предыдущих
                                   join t in resultTek
                                   on new { OKPO = p.OKPO, OKVED = p.OKVED } equals new { OKPO = t.OKPO, OKVED = t.OKVED }
                                   into data_A
                                   from data_B in data_A.DefaultIfEmpty(new Model())
                                   select new ModelResult()
                                   {
                                       OKPO = p.OKPO,
                                       OKVEDPred = p.OKVED,
                                       OKVEDTek = data_B.OKVED
                                   };

                var queryPred = leftJoinPred.Where(z => string.IsNullOrEmpty(z.OKVEDTek));

                var leftJoinTek = from t in resultTek // left join для текущих
                                  join p in resultPred
                                  on new { OKPO = t.OKPO, OKVED = t.OKVED } equals new { OKPO = p.OKPO, OKVED = p.OKVED }
                                  into data_A
                                  from data_B in data_A.DefaultIfEmpty(new Model())
                                  select new ModelResult()
                                  {
                                      OKPO = t.OKPO,
                                      OKVEDTek = t.OKVED,
                                      OKVEDPred = data_B.OKVED
                                  };
                
                var queryTek = leftJoinTek.Where(z => string.IsNullOrEmpty(z.OKVEDPred));


                var joinPredTek = leftJoinTek.Join(leftJoinPred, t => t.OKPO, p => p.OKPO, (t,p)=> new ModelResult {OKPO = t.OKPO, OKVEDTek = t.OKVEDTek, OKVEDPred = p.OKVEDPred}).Where(z=>z.OKVEDPred!=z.OKVEDTek);
                
                var resultOuterJoin =queryPred.Union(queryTek).Union(joinPredTek);


                //var resultCompareOKVED = from q1 in queryPred
                //                         join q2 in queryTek
                //                         on q1.OKPO equals q2.OKPO
                //                         select new ModelResult() { OKPO = q1.OKPO, OKVEDPred = q1.OKVEDPred, OKVEDTek = q2.OKVEDTek };

                //var okpoPredNotInOkpoTek = from q1 in queryPred
                //                           from q2 in queryTek
                //                           where q1.OKPO != q2.OKPO
                //                           select new ModelResult() { OKPO = q1.OKPO, OKVEDPred = q1.OKVEDPred, OKVEDTek = q1.OKVEDTek };

                ////listOkpoPredNotInOkpoTek = okpoPredNotInOkpoTek.ToList();


                //var okpoTekNotInOkpoPred = from q1 in queryTek
                //                           from q2 in queryPred
                //                           where q1.OKPO != q2.OKPO
                //                           select new ModelResult() { OKPO = q1.OKPO, OKVEDPred = q1.OKVEDPred, OKVEDTek = q1.OKVEDTek };

                //var a = okpoPredNotInOkpoTek.GroupBy(z => new { z.OKPO, z.OKVEDTek }).Select(z => z.First()).ToList();
                //var b = okpoTekNotInOkpoPred.GroupBy(z => new { z.OKPO, z.OKVEDPred }).Select(z => z.First()).ToList();


                //result = a.Concat(b).ToList();

                //result = (List<Model>)okpoPredNotInOkpoTek.GroupBy(z => new { z.OKPO, z.OKVED }).Select(z => z.First()).Concat(okpoTekNotInOkpoPred.GroupBy(z => new { z.OKPO, z.OKVED}).Select(z => z.First()));

                return resultOuterJoin.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                //result = null;
                return null;
            }
        }
    }
}
