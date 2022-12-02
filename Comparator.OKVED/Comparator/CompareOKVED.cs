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
            var delokvedAG = DelOkvedAG(collectionPBD).GroupBy(a => new { a.OKPO, a.OKVEDHoz, a.OKVEDChist }).Select(a => a.First());
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
