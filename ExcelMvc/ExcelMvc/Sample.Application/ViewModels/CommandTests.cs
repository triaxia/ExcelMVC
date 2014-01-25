using System.Windows;
using ExcelMvc.Controls;
using ExcelMvc.Views;

namespace Sample.Application.ViewModels
{
    internal class CommandTests
    {
        private Sheet View { get; set; }
        public CommandTests(Sheet sheet)
        {
            View = sheet;
            sheet.HookClicked(CmdClicked, "ShapeButton", true);
            sheet.HookClicked(CmdClicked, "FormButton", true);
            sheet.HookClicked(CmdClicked, "FormCheckBox", true);
            sheet.HookClicked(CmdClicked, "FormOptionButtonYes", true);
            sheet.HookClicked(CmdClicked, "FormOptionButtonNo", true);
            sheet.HookClicked(CmdClicked, "FormOptionButtonMale", true);
            sheet.HookClicked(CmdClicked, "FormOptionButtonFemale", true);
            sheet.HookClicked(CmdClicked, "FormListBox", true);
            sheet.HookClicked(CmdClicked, "FormDropDown", true);
            sheet.HookClicked(CmdClicked, "FormSpinner", true);
        }

        public void CmdClicked(object sender, CommandEventArgs args)
        {
            var cmd = (Command) sender;
            var message = string.Format("Command (name={0}, caption={1}, value={2}, enabled={3}) clicked.",
                cmd.Name, cmd.Caption, cmd.Value, cmd.IsEnabled);
            MessageBox.Show(message, View.Name, MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
