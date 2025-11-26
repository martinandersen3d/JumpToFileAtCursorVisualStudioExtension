using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell.Interop;

namespace JumpToFileAtCursorVisualStudioExtension.Commands;

[Command(PackageIds.JumpToFileCommand)]
internal sealed class JumpToFileCommand : BaseCommand<JumpToFileCommand>
{
    protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        try
        {
            // 1. Get the DTE (Development Tools Environment) service
            DTE2 dte = await VS.GetServiceAsync<DTE, DTE2>();

            if (dte?.ActiveDocument == null)
            {
                // No file is currently open
                return;
            }

            // 2. Get the text selection
            TextSelection selection = (TextSelection)dte.ActiveDocument.Selection;
            TextPoint point = selection.ActivePoint;

            // If nothing is explicitly selected, expand selection to the word under cursor
            if (selection.IsEmpty)
            {
                // Expand selection to the current word using WordLeft/WordRight
                selection.WordLeft(true);
                selection.WordRight(true);
            }

            string searchText = selection.Text.Trim();

            if (string.IsNullOrEmpty(searchText))
            {
                return;
            }

            // 3. Find the file (Helper method call)
            // We use a helper because Solutions can have nested folders
            ProjectItem foundItem = FindFileInSolution(dte.Solution, searchText);

            if (foundItem != null)
            {
                // 4. Open the file
                Window window = foundItem.Open();
                window.Activate();
            }
            else
            {
                await VS.MessageBox.ShowAsync(
                    $"Could not find a file named '{searchText}'",
                    "File Not Found",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK);
            }
        }
        catch (Exception ex)
        {
            // Log error to the Output window
            await ex.LogAsync();
        }
    }

    // Helper method to recursively search for the file
    private ProjectItem FindFileInSolution(EnvDTE.Solution solution, string fileName)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        foreach (EnvDTE.Project project in solution.Projects)
        {
            ProjectItem item = FindItemInProject(project.ProjectItems, fileName);
            if (item != null) return item;
        }
        return null;
    }

    private ProjectItem FindItemInProject(ProjectItems items, string fileName)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (items == null) return null;

        foreach (ProjectItem item in items)
        {
            // Check if the item name matches (ignoring extension for smarter matching)
            // You might want to adjust this to 'Equals' if you want exact matches.
            if (item.Name.StartsWith(fileName, StringComparison.OrdinalIgnoreCase))
            {
                return item;
            }

            // Recursion for folders/sub-items
            if (item.ProjectItems != null && item.ProjectItems.Count > 0)
            {
                ProjectItem foundSub = FindItemInProject(item.ProjectItems, fileName);
                if (foundSub != null) return foundSub;
            }
        }
        return null;
    }
}
