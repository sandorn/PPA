using NetOffice.VBIDEApi.Enums;
using Project.Utilities;
using System.Diagnostics;
using NETOP = NetOffice.PowerPointApi;

namespace VBAApi
{
	public static class VbaExecutor
	{
		#region Public Methods

		public static void ExecuteVbaCode(NETOP.Application app,string vbaCode,string macroName,params object[] args)
		{
			dynamic presentation = app.ActivePresentation;

			ExHandler.Run(() =>
			{
				var vbProject = presentation.VBProject;
				var module = vbProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
				var codeModule = module.CodeModule;
				codeModule.AddFromString(vbaCode);
				presentation.Save();    //触发编译
				Debug.WriteLine($"模块创建：{module.Name}({codeModule.CountOfLines} lines) | args.Length  : {args.Length} ");

				/*—运行宏—*/
				if(args.Length == 0) app.Run($"{module.Name}.{macroName}");
				else app.Run($"{module.Name}.{macroName}",args);

				/*—清理—*/
				//短暂延迟防锁COM对象
				System.Threading.Thread.Sleep(50);
				vbProject.VBComponents.Remove(module);
			},"ExecuteVbaCode");
		}

		#endregion Public Methods
	}
}