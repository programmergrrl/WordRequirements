using ProgrammerGrrl.WordRequirements.Core;

namespace ProgrammerGrrl.WordRequirements.Shim2007
{
    public partial class AddInModule
    {
        private void AddInModule_Startup(object sender, System.EventArgs e)
        {
            ApplicationContext.Application = Application;
        }

        private void AddInModule_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(AddInModule_Startup);
            this.Shutdown += new System.EventHandler(AddInModule_Shutdown);
        }
        
        #endregion
    }
}
