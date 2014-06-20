
namespace SpotTrading.ViewModels
{
    using System.Linq;
    using BusinessModels;
    using CmdModels;
    using ExcelMvc.Controls;
    using ExcelMvc.Views;

    public class ViewModelTrading
    {
        public ViewModelTrading(View book)
        {
            // static ccy pair table (OneWayToSource)
            var tblCcyPair = (Table)book.Find("ExcelMvc.Table.CcyPairs");
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
            book.FindCommand("ExcelMvc.Command.InsideMode").Clicked += (x, y) =>
            {
                deal.IsInsideTrading = System.Convert.ToBoolean(((Command)x).Value);
            };

            // position table
            var positions = new ViewModelPositions();
            book.Find("ExcelMvc.Table.Positions").Model = positions;
            book.FindCommand("ExcelMvc.Command.Reset").Clicked += (x, y) => positions.Reset();

            // manual deal command
            book.FindCommand("ExcelMvc.Command.ManualDeal").Model = new CmdModelManualDeal(deal, positions, rates);

            var dealing = new ViewModelDealing(pairs.Ccys.ToList(), deal, positions, rates);
            book.FindCommand("ExcelMvc.Command.AutoDeal").Model = new CmdModelAutoDeal(dealing);
        }
    }
}
