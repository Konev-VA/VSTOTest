using System.ComponentModel;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;

namespace B2Green
{
	/// <summary>
	/// ViewModel для кастомной вкладки
	/// </summary>
	public class RibbonViewModel : INotifyPropertyChanged
	{
		public ICommand Command { get; }

		public event PropertyChangedEventHandler PropertyChanged;

		public RibbonViewModel()
		{
			Command = new Command(
				execute: x => TurnB2CellToGreen(),
				canExecute: _ => true
			);
		}

		/// <summary>
		/// Покрасить ячейку B2 в зелёный цвет
		/// </summary>
		public void TurnB2CellToGreen()
		{
			Excel.Worksheet worksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
			Excel.Range cell = worksheet.get_Range("B2");
			cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
		}
	}
}
