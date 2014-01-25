using System.Windows;
using ExcelMvc.Controls;
using ExcelMvc.Views;

namespace Sample.Application.ViewModels
{
    internal class Tests
    {
        public Tests(Sheet sheet)
        {
            sheet.HookClicked(CmdClicked, "ShapeButton", true);
            sheet.HookClicked(CmdClicked, "FormButton", true);
            sheet.HookClicked(CmdClicked, "FormCheckBox", true);
            sheet.HookClicked(CmdClicked, "FormOptionButtonYes", true);
            sheet.HookClicked(CmdClicked, "FormOptionButtonNo", true);
        }

        public void CmdClicked(object sender, CommandEventArgs args)
        {
            var cmd = (Command) sender;
            MessageBox.Show(string.Format("Command (name={0}, caption={1}, value={2}, enabled={3}) clicked.",
                cmd.Name, cmd.Caption, cmd.Value, cmd.IsEnabled));
        }
    }
}
