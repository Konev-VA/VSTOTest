using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace B2Green
{
	[ComVisible(true)]
	public class Ribbon : Office.IRibbonExtensibility
	{
		private Office.IRibbonUI ribbon;

		private RibbonViewModel _viewModel;

		private RibbonViewModel ViewModel
		{
			get
			{
				if (_viewModel == null)
				{
					_viewModel = new RibbonViewModel();
				}
				return _viewModel;
			}
		}

		public void OnButtonClick(Office.IRibbonControl control)
		{
			ViewModel.Command.Execute(null);
		}

		public bool IsButtonEnabled(Office.IRibbonControl control)
		{
			return ViewModel.Command.CanExecute(null);
		}

		#region Элементы IRibbonExtensibility

		public string GetCustomUI(string ribbonID)
		{
			return GetResourceText("B2Green.Ribbon1.xml");
		}

		#endregion

		#region Обратные вызовы ленты
		//Информацию о методах создания обратного вызова см. здесь. Дополнительные сведения о методах добавления обратного вызова см. по ссылке https://go.microsoft.com/fwlink/?LinkID=271226

		public void Ribbon_Load(Office.IRibbonUI ribbonUI)
		{
			this.ribbon = ribbonUI;
		}

		#endregion

		#region Вспомогательные методы

		private static string GetResourceText(string resourceName)
		{
			Assembly asm = Assembly.GetExecutingAssembly();
			string[] resourceNames = asm.GetManifestResourceNames();
			for (int i = 0; i < resourceNames.Length; ++i)
			{
				if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
				{
					using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
					{
						if (resourceReader != null)
						{
							return resourceReader.ReadToEnd();
						}
					}
				}
			}
			return null;
		}

		#endregion
	}
}
