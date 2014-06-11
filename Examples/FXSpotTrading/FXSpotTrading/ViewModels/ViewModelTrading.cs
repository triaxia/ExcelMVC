using ExcelMvc.Controls;
using ExcelMvc.Views;
using FXSpotTrading.Models;

namespace FXSpotTrading.ViewModels
{
    public class ViewModelTrading
    {
        private ViewModelExchangeRates ExchangeRates { get; set; }
        public ViewModelTrading(View book)
        {
            // bind static ccy pair table (OneWayToSource)
            var tblCcyPair = (Table) book.Find("ExcelMvc.Table.CcyPairs");
            var pairs = new CcyPairs(tblCcyPair.MaxItemsToBind);
            tblCcyPair.Model = pairs;

            // bind static ccy list (OneWay)
            var tblCcys = (Table)book.Find("ExcelMvc.Table.Ccys");
            tblCcys.Model = pairs.Ccys;

            // bind exchange rates
            var tblRates = (Table)book.Find("ExcelMvc.Table.Rates");
            tblRates.Model = ExchangeRates = new ViewModelExchangeRates(new ExchangeRates(pairs));

            book.FindCommand("AutoRate").Model = ExchangeRates;
        }
    }
}
