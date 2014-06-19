using ExcelMvc.Views;

namespace FXSpotTrading.ViewModels
{
    using BusinessModels;
    using CmdModels;

    public class ViewModelTrading
    {
        public ViewModelTrading(View book)
        {
            // static ccy pair table (OneWayToSource)
            var tblCcyPair = (Table) book.Find("ExcelMvc.Table.CcyPairs");
            var pairs = new CcyPairs(tblCcyPair.MaxItemsToBind);
            tblCcyPair.Model = pairs;

            // static ccy list (OneWay)
            var tblCcys = book.Find("ExcelMvc.Table.Ccys");
            tblCcys.Model = pairs.Ccys;

            // exchange rates
            var tblRates = book.Find("ExcelMvc.Table.Rates");
            var rates = new ViewModelExchangeRates(new ExchangeRates(pairs));
            tblRates.Model = rates;

            // auto rate command
            book.FindCommand("ExcelMvc.Command.AutoRate").Model = new CmdModelAutoRate(rates);

            // deal form
            var deal = new ViewModelDeal(rates);
            book.Find("ExcelMvc.Form.Deal").Model = deal;

            // manual deal command
            book.FindCommand("ExcelMvc.Command.ManualDeal").Model = new CmdModelDeal(deal);
        }
    }
}
