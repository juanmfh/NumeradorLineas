using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Text.Tagging;
using Microsoft.VisualStudio.Text.Classification;
using System.Drawing.Imaging;

namespace NumeradorLineas
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Numerador 
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("7f69ccf1-02ce-4327-b778-3835f0f2a5c9");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Numerador"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private Numerador(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Numerador Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in Numerador's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new Numerador(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            NumerarLineas();
        }

        private void NumerarLineas()
        {
            // Obtiene la instancia del editor de texto actual
            IWpfTextView textView = GetTextView();

            if (textView != null)
            {
                // Obtiene el snapshot completo del texto
                ITextSnapshot snapshot = textView.TextBuffer.CurrentSnapshot;

                // Obtiene la selección actual
                ITextSnapshotLine startLine = textView.TextSnapshot.GetLineFromPosition(textView.Selection.Start.Position);
                ITextSnapshotLine endLine = textView.TextSnapshot.GetLineFromPosition(textView.Selection.End.Position);

                int cont = 100;

                bool ifAbierto = false;
                bool selectCaseAbierto = false;

                // Recorre cada línea y agrega el número
                for (int lineNumber = startLine.LineNumber; lineNumber <= endLine.LineNumber; lineNumber++)
                {
                    snapshot = textView.TextBuffer.CurrentSnapshot;

                    ITextSnapshotLine line = snapshot.GetLineFromLineNumber(lineNumber);
                    string lineText = line.GetText();

                    if (!string.IsNullOrEmpty(lineText))
                    {
                        int currentNum = 0;
                        int iPos = line.Start.Position;

                        // Reemplaza el texto de la línea con el número
                        using (ITextEdit edit = snapshot.TextBuffer.CreateEdit())
                        {
                            // Si la línea ya tiene contador lo elimino e inserto el nuevo
                            if (lineText.Contains(":") && int.TryParse(lineText.Substring(0, lineText.IndexOf(":")), out currentNum))
                            {
                                edit.Delete(iPos, lineText.IndexOf(":") + 1);
                                edit.Apply();
                            }
                        }

                        bool bIgnorarNumeracion = false;

                        if (selectCaseAbierto)
                        {
                            if (firstWord(lineText.ToUpper()).StartsWith("CASE"))
                            {
                                selectCaseAbierto = false;

                                bIgnorarNumeracion = true;
                            }
                        }

                        if(ifAbierto)
                        {
                            if (lineText.EndsWith("Then"))
                            {
                                ifAbierto = false;

                                bIgnorarNumeracion = true;
                            }
                        }

                        if (!bIgnorarNumeracion)
                        {
                            if (firstWord(lineText.ToUpper()).StartsWith("SELECT"))
                            {
                                selectCaseAbierto = true;
                            }
                            else if (firstWord(lineText.ToUpper()).StartsWith("IF") && !lineText.ToUpper().EndsWith("Then"))
                            {
                                // Si hay un if multilínea no se numera la línea hasta que termine
                                ifAbierto = true;
                            }

                            // Reemplaza el texto de la línea con el número
                            using (ITextEdit edit = snapshot.TextBuffer.CreateEdit())
                            {
                                edit.Insert(iPos, cont.ToString() + ": ");
                                cont += 10;
                                edit.Apply();
                            }
                        }
                    }
                }
            }
        }


        private String firstWord(String line)
        {
            int currentNum;
            string linproc = line;
            if (line.Contains(":") && int.TryParse(line.Substring(0, line.IndexOf(":")), out currentNum))
            {
                linproc = line.Substring(line.IndexOf(":") + 1);
            }

            int indexSpace = linproc.Trim().IndexOf(' ');
            int indexTab = linproc.Trim().IndexOf("\t");

            if (indexSpace != -1 && indexTab != -1 )
            {
                if (indexSpace < indexTab)
                {
                    return linproc.Trim().Substring(0, indexSpace);
                }
                else
                {
                    return linproc.Trim().Substring(0, indexTab);
                }
            }
            else if (indexSpace != -1)
            {
                return linproc.Trim().Substring(0, indexSpace);
            }
            else if (indexTab != -1)
            {
                return linproc.Trim().Substring(0, indexTab);
            }
            else
            {
                return linproc.Trim();
            }
        }

        private IWpfTextView GetTextView()
        {
            // Obtiene la instancia de IVsTextView mediante el servicio de texto
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            IVsTextManager textManager = (Package.GetGlobalService(typeof(SVsTextManager)) as IVsTextManager);

            IVsTextView textView;
            textManager.GetActiveView(1, null, out textView);

            // Convierte a IWpfTextView
            IVsUserData userData = textView as IVsUserData;
            IWpfTextViewHost textViewHost = null;

            if (userData != null)
            {
                Guid guidViewHost = DefGuidList.guidIWpfTextViewHost;
                userData.GetData(ref guidViewHost, out var objTextViewHost);
                textViewHost = (IWpfTextViewHost)objTextViewHost;
            }

            return textViewHost?.TextView;
        }
    }
}
