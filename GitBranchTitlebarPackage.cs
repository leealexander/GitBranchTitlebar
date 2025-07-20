using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Diagnostics;
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
        private string _currentBranch = string.Empty;
        private string _lastForcedTitle = string.Empty;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await base.InitializeAsync(cancellationToken, progress);

            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            var dte = await GetServiceAsync(typeof(DTE)) as DTE2;

            if (dte != null)
            {
                _dte = dte;
                _timer = new Timer(TimerCallback, null, 0, 5000);
            }
        }

        private async Task UpdateReferencesAsync(object state)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var solutionName = GetSolutionName();
            if(solutionName is null)
            {
                return;
            }

            string branchName = GetCurrentGitBranch();

            if (!string.IsNullOrEmpty(branchName))
            {
                var newTitle = $"{solutionName} - [{branchName}]";
                var currentTitle = Win32TitleHelper.GetVSMainWindowTitle(_dte.MainWindow.Caption);
                if (string.IsNullOrEmpty(currentTitle))
                {
                    currentTitle = _lastForcedTitle;
                }
                if (currentTitle != newTitle)
                {
                    Win32TitleHelper.SetVSMainWindowTitle(currentTitle, newTitle);
                    _lastForcedTitle = newTitle;
                }

                if (_currentBranch != branchName)
                {
                    _currentBranch = branchName;
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
        }

        private string GetSolutionName()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_dte.Solution == null || string.IsNullOrEmpty(_dte.Solution.FullName))
                return null;

            // Extract just the file name without its extension
            return Path.GetFileNameWithoutExtension(_dte.Solution.FullName);
        }

        private async void TimerCallback(object state)
        {
            await ThreadHelper.JoinableTaskFactory.RunAsync(() => UpdateReferencesAsync(state));
        }

        private string GetCurrentGitBranch()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_dte.Solution == null || string.IsNullOrEmpty(_dte.Solution.FullName))
                return string.Empty;

            string solutionDir = Path.GetDirectoryName(_dte.Solution.FullName);
            return GetGitBranchFromDirectory(solutionDir);
        }

        private string GetGitBranchFromDirectory(string directory)
        {
            try
            {
                // Walk up the directory tree to find .git directory
                string currentDir = directory;
                while (!string.IsNullOrEmpty(currentDir))
                {
                    string gitDir = Path.Combine(currentDir, ".git");
                    
                    if (Directory.Exists(gitDir))
                    {
                        // Found .git directory, try to read HEAD file
                        string headFile = Path.Combine(gitDir, "HEAD");
                        if (File.Exists(headFile))
                        {
                            string headContent = File.ReadAllText(headFile).Trim();
                            
                            if (headContent.StartsWith("ref: refs/heads/"))
                            {
                                // Extract branch name from ref
                                return headContent.Substring("ref: refs/heads/".Length);
                            }
                            else if (headContent.Length >= 7)
                            {
                                // Detached HEAD - return short commit hash
                                return headContent.Substring(0, 7);
                            }
                        }
                        break;
                    }
                    else if (File.Exists(gitDir))
                    {
                        // .git file (git worktree)
                        string gitFileContent = File.ReadAllText(gitDir).Trim();
                        if (gitFileContent.StartsWith("gitdir: "))
                        {
                            string realGitDir = gitFileContent.Substring("gitdir: ".Length);
                            if (!Path.IsPathRooted(realGitDir))
                            {
                                realGitDir = Path.Combine(currentDir, realGitDir);
                            }
                            
                            string headFile = Path.Combine(realGitDir, "HEAD");
                            if (File.Exists(headFile))
                            {
                                string headContent = File.ReadAllText(headFile).Trim();
                                
                                if (headContent.StartsWith("ref: refs/heads/"))
                                {
                                    return headContent.Substring("ref: refs/heads/".Length);
                                }
                                else if (headContent.Length >= 7)
                                {
                                    return headContent.Substring(0, 7);
                                }
                            }
                        }
                        break;
                    }
                    
                    // Move up one directory level
                    string parentDir = Path.GetDirectoryName(currentDir);
                    if (parentDir == currentDir) // Reached root
                        break;
                    currentDir = parentDir;
                }
            }
            catch (Exception)
            {
                // Ignore exceptions and fall back to git command
            }

            // Fallback: try to use git command
            return GetGitBranchFromCommand(directory);
        }

        private string GetGitBranchFromCommand(string workingDirectory)
        {
            try
            {
                var processInfo = new ProcessStartInfo
                {
                    FileName = "git",
                    Arguments = "rev-parse --abbrev-ref HEAD",
                    WorkingDirectory = workingDirectory,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var process = System.Diagnostics.Process.Start(processInfo))
                {
                    if (process != null)
                    {
                        process.WaitForExit(5000); // 5 second timeout
                        if (process.ExitCode == 0)
                        {
                            string output = process.StandardOutput.ReadToEnd().Trim();
                            if (!string.IsNullOrEmpty(output) && output != "HEAD")
                            {
                                return output;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Git command failed, ignore
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

