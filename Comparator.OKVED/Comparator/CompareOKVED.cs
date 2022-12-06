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
                KodPokaz = a.KodPokaz,
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
                                     KodPokaz = cur.KodPokaz,
                                     OKATO = cur.OKATO,
                                     OKVEDCurPer = cur.OKVEDChist,
                                     OKVEDPrevPer = data_B.OKVEDChist
                                 };

            var notInPrevPer = leftJoinCurPer.Where(a => string.IsNullOrEmpty(a.OKVEDPrevPer));

            var leftJoinPrevPer = from prev in groupOkpoOkvedPrevPer  //left join для текущего периода
                                  join cur in groupOkpoOkvedCurPer
                                 on new { OKPO = prev.OKPO, OKVED = prev.OKVEDChist } equals new { OKPO = cur.OKPO, OKVED = cur.OKVEDChist }
                                 into data_A
                                 from data_B in data_A.DefaultIfEmpty(new ModelPBD())
                                 select new ModelResultChist()
                                 {
                                     Period = prev.Period,
                                     OKPO = prev.OKPO,
                                     Name = prev.Name,
                                     KodPokaz = prev.KodPokaz,
                                     OKATO = prev.OKATO,
                                     OKVEDCurPer = data_B.OKVEDChist,
                                     OKVEDPrevPer = prev.OKVEDChist
                                 };

            var notInCurPer = leftJoinPrevPer.Where(a => string.IsNullOrEmpty(a.OKVEDCurPer));

            //var periodMes = delAGOkved.Where(a => a.Period == "Нет");
            //var otchMesNotNull = periodMes.Where(a => a.OtchMes != null && a.PredMes == null);
            //var predMesNotNull = periodMes.Where(a => a.PredMes != null && a.OtchMes == null);
            //var resultCompareMes = otchMesNotNull.Join(predMesNotNull, a => new { a.OKPO, a.KodPokaz }, b => new { b.OKPO, b.KodPokaz },
            //    (a, b) => new ModelResultChist { Period = a.Period, OKPO = a.OKPO, Name = a.Name, OKATO = a.OKATO, KodPokaz = a.KodPokaz, ChistOkvedOtchMes = a.OKVEDChist, ChistOKVEDPredMes = b.OKVEDChist });

            //var periodKvart = delAGOkved.Where(a => a.Period == "Да");
            //var otchKvartNotNull = periodKvart.Where(a => a.OtchKvart != null && a.PredKvart == null);
            //var predKvartNotNull = periodKvart.Where(a => a.PredKvart != null && a.OtchKvart == null);
            //var resultCompareKvart = otchKvartNotNull.Join(predKvartNotNull, a => new { a.OKPO, a.KodPokaz }, b => new { b.OKPO, b.KodPokaz },
            //    (a, b) => new ModelResultChist { Period = a.Period, OKPO = a.OKPO, Name = a.Name, OKATO = a.OKATO, KodPokaz = a.KodPokaz, ChistOkvedOtchKvart=a.OKVEDChist, ChistOKVEDPredKvart = b.OKVEDChist });

            //var periodMes = delAGOkved.Where(a => a.Period == "Нет");
            //var resultMes = periodMes.Where(a => a.OtchMes != null && a.PredMes == null && a.PerSnachOtchGod == null && a.SovMesPredGod == null && a.SovPerPredGod == null &&
            //a.OtchKvart == null && a.PredKvart == null && a.SovKvartPredGod == null).Select(a => new ModelResultChist { Period = a.Period, OKPO = a.OKPO, Name = a.Name, OKATO = a.OKATO, KodPokaz = a.KodPokaz, ChistOkvedOtchMes = a.OKVEDChist });

            //var periodKvart = delAGOkved.Where(a => a.Period == "Нет");
            //var resultKvart = periodKvart.Where(a => a.OtchKvart != null && a.PredMes == null && a.PerSnachOtchGod == null && a.SovMesPredGod == null && a.SovPerPredGod == null &&
            //a.OtchKvart == null && a.PredKvart == null && a.SovKvartPredGod == null).Select(a => new ModelResultChist { Period = a.Period, OKPO = a.OKPO, Name = a.Name, OKATO = a.OKATO, KodPokaz = a.KodPokaz, ChistOkvedOtchMes = a.OKVEDChist });


            return notInPrevPer.Union(notInCurPer);
        }
        private IEnumerable<ModelPBD> GetRowsOkvedHozEqualsChist(IEnumerable<ModelPBD> collectionPBD) => collectionPBD.Where(a => a.OKVEDHoz == a.OKVEDChist);
        private IEnumerable<ModelPBD> DelOkvedAG(IEnumerable<ModelPBD> collectionPBD) => collectionPBD.Where(a => a.OKVEDChist != "101.АГ");
        private PropertyInfo GetPropertyInfoRadioButton(object radioButtonOrderBy)
        {
            string paramName = string.Empty;

            var modelResultHozChist = new ModelResultHozChist();
            switch (radioButtonOrderBy)
            {
                case "ОКПО":
                    paramName = nameof(modelResultHozChist.OKPO);
                    break;
                case "Наименованию предприятия":
                    paramName = nameof(modelResultHozChist.Name);
                    break;
                case "ОКАТО":
                    paramName = nameof(modelResultHozChist.OKATO);
                    break;
                case "Хозяйственному ОКВЭД":
                    paramName = nameof(modelResultHozChist.OKVEDHoz);
                    break;
                case "Чистому ОКВЭД":
                    paramName = nameof(modelResultHozChist.OKVEDChist);
                    break;
                default:
                    break;
            }

            return typeof(ModelResultHozChist).GetProperty(paramName);
        }
    }
}
