using ExcelMvc.Views;
using FXSpotTrading.Models;

namespace FXSpotTrading.ViewModels
{
    public class ViewModelTrading
    {
        public ViewModelTrading(View book)
        {
            // bind static ccy pair table (OneWayToSource)
            var tblCcyPair = (Table) book.Find("ExcelMvc.Table.CcyPairs");
            var pairs = new CcyPairs(tblCcyPair.MaxItemsToBind);
            tblCcyPair.Model = pairs;

            // bind static ccy list (OneWay)
            var tblCcys = book.Find("ExcelMvc.Table.Ccys");
            tblCcys.Model = pairs.Ccys;

            // bind exchange rates
            var tblRates = book.Find("ExcelMvc.Table.Rates");
            var rates = new ViewModelExchangeRates(new ExchangeRates(pairs));
            tblRates.Model = rates;

            book.FindCommand("ExcelMvc.AutoRate").Model = rates;

            // bind deal
            book.Find("ExcelMvc.Form.Deal").Model = new ViewModelDeal(rates);
        }
    }
}
