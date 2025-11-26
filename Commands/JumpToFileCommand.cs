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

            string searchText;

            // If nothing is explicitly selected, compute the word under the cursor manually
            if (selection.IsEmpty)
            {
                searchText = GetWordUnderCursor(point);
            }
            else
            {
                searchText = selection.Text.Trim();
            }

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

    // Get the full word under the cursor. Word characters allowed: letters, digits, '-' and '_'.

    private string GetWordUnderCursor(TextPoint point)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var editPoint = point.CreateEditPoint();
        string line = editPoint.GetLines(point.Line, point.Line + 1);

        if (string.IsNullOrEmpty(line)) return string.Empty;

        int index = Math.Max(0, point.LineCharOffset - 1);
        if (index >= line.Length) index = line.Length - 1;

        // If the caret is on a non-word character, try to move right to find a word
        if (!IsWordChar(line[index]))
        {
            int temp = index;
            while (temp < line.Length && !IsWordChar(line[temp])) temp++;
            if (temp >= line.Length) return string.Empty;
            index = temp;
        }

        int left = index;
        while (left > 0 && IsWordChar(line[left - 1])) left--;

        int right = index;
        while (right < line.Length && IsWordChar(line[right])) right++;

        int len = right - left;
        if (len <= 0) return string.Empty;

        return line.Substring(left, len);
    }

    private bool IsWordChar(char c)
    {
        return char.IsLetterOrDigit(c) || c == '-' || c == '_';
    }
}
