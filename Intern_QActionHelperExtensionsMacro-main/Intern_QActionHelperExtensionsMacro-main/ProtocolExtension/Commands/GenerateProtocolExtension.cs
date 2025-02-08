// <copyright file="GenerateProtocolExtension.cs" company="Skyline Communications">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace ProtocolExtension.Commands
{
    using System;
    using System.ComponentModel.Design;
    using System.Windows.Forms;
    using EnvDTE;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell;
    using Task = System.Threading.Tasks.Task;

    /// <summary>
    /// Command handler.
    /// </summary>
    internal sealed class GenerateProtocolExtension
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("8e96ce60-db99-451b-8bac-624168c3232a");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="GenerateProtocolExtension"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file).
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private GenerateProtocolExtension(AsyncPackage package, OleMenuCommandService commandService)
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
        public static GenerateProtocolExtension Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in Command1's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new GenerateProtocolExtension(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        [STAThread]
        private void Execute(object sender, EventArgs e)
        {
            // Generate the code
            var code = ProtocolExtensionGenerator.CreateMethods();

            try
            {
                // Get the currently active instance of Visual Studio (which is the experimental instance in this case)
                var dte = (DTE2)System.Runtime.InteropServices.Marshal.GetActiveObject("VisualStudio.DTE");

                if (dte != null)
                {
                    // Get the active document in the experimental Visual Studio
                    var activeDocument = dte.ActiveDocument;

                    if (activeDocument != null)
                    {
                        // Get the active text document (the file that is open in the editor)
                        var textDocument = (TextDocument)activeDocument.Object("TextDocument");

                        // Create an edit point at the end of the document
                        var editPoint = textDocument.CreateEditPoint();
                        editPoint.EndOfDocument();

                        // Insert the generated code at the end
                        editPoint.Insert(code);

                        // Optionally, copy the generated code to the clipboard
                        Clipboard.SetText(code);

                        // Show a message box indicating success
                        MessageBox.Show("Methods successfully generated and inserted into the current document!");
                    }
                    else
                    {
                        MessageBox.Show("No active document found!");
                    }
                }
                else
                {
                    MessageBox.Show("Could not retrieve the Visual Studio environment!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
        }
    }
}
