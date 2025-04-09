using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TeamFoundation.Git.Extensibility;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace GitBranchTitlebar
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("Git Branch Titlebar", "Shows current Git branch in window title", "1.0")]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(GitBranchTitlebarPackage.PackageGuidString)]
    public sealed class GitBranchTitlebarPackage : AsyncPackage
    {
        public const string PackageGuidString = "2bec722a-4155-4b1f-af6c-6bb9f5474803";
        private DTE2 _dte;
        private Timer _timer;
        private const string BranchTitleFormat = "{0} - [{1}]";
        private string _currentBranchName = string.Empty;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await base.InitializeAsync(cancellationToken, progress);

            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            var dte = await GetServiceAsync(typeof(DTE)) as DTE2;

            if (dte != null)
            {
                _dte = dte;
                _timer = new Timer(UpdateWindowTitle, null, 0, 5000);
            }
        }




        private async Task UpdateWindowTitleAsync(object state)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var solutionName = GetSolutionName();
            if(solutionName is null)
            {
                return;
            }

            string branchName = GetCurrentGitBranch();

            if (!string.IsNullOrEmpty(branchName) && _currentBranchName != branchName)
            {
                _currentBranchName = branchName;
                Win32TitleHelper.SetVSMainWindowTitle(_dte, $"{solutionName} - [{branchName}]");
                var solutionPath = _dte.Solution.FullName;
                var staThread = new System.Threading.Thread(() =>
                {
                    JumpListHelper.UpdateSolutionJumpListEntry(solutionPath, branchName);
                });

                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start();
                staThread.Join();
            }
        }

        private string GetSolutionName()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_dte.Solution == null || string.IsNullOrEmpty(_dte.Solution.FullName))
                return null;

            // Extract just the file name without its extension
            return Path.GetFileNameWithoutExtension(_dte.Solution.FullName);
        }

        private async void UpdateWindowTitle(object state)
        {
            await ThreadHelper.JoinableTaskFactory.RunAsync(() => UpdateWindowTitleAsync(state));
        }

        private string GetCurrentGitBranch()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            IGitExt gitExt = ServiceProvider.GlobalProvider.GetService(typeof(IGitExt)) as IGitExt;

            if (gitExt == null || gitExt.ActiveRepositories == null)
                return string.Empty;

            foreach (IGitRepositoryInfo repo in gitExt.ActiveRepositories)
            {
                if (repo == null)
                    continue;

                // Check for regular branch
                if (!string.IsNullOrEmpty(repo.CurrentBranch?.Name))
                    return repo.CurrentBranch.Name;

                // Handle Git WorkTree (detached HEAD)
                if (!string.IsNullOrEmpty(repo.RepositoryPath))
                {
                    string headFilePath = Path.Combine(repo.RepositoryPath, "HEAD");
                    if (File.Exists(headFilePath))
                    {
                        string headContent = File.ReadAllText(headFilePath).Trim();

                        if (headContent.StartsWith("ref: "))
                        {
                            string branchRef = headContent.Substring(5).Trim();
                            return branchRef.Replace("refs/heads/", string.Empty);
                        }
                        else
                        {
                            return headContent.Substring(0, 7); // return commit hash prefix if truly detached
                        }
                    }
                }
            }

            return string.Empty;
        }

        protected override void Dispose(bool disposing)
        {
            _timer?.Dispose();
            base.Dispose(disposing);
        }

    }
}

