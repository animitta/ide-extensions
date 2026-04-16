using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace VsCheckerExtension
{
  [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
  [InstalledProductRegistration("VsCheckerExtension", "C# 文件校验扩展", "1.1")]
  [ProvideMenuResource("Menus.ctmenu", 1)]
  [ProvideOptionPage(typeof(VsCheckerOptions), "VsChecker", "文件已校验", 0, 0, true)]
  [ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
  [Guid("A59B663D-9E90-46EF-933E-6C2AF9C5FE4E")]
  public sealed class VsCheckerPackage : AsyncPackage
  {
    protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
    {
      await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
      await VsCheckerCommand.InitializeAsync(this);
    }
  }
}
