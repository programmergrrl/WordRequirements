using Microsoft.Office.Tools.Ribbon;
using ProgrammerGrrl.WordRequirements.Core.Schema;

namespace ProgrammerGrrl.WordRequirements.Shim2007
{
    public partial class RequireRibbon
    {
        private void RequireRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            new SchemaAdder().AddSchema();
        }
    }
}
