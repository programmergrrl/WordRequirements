using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace ProgrammerGrrl.WordRequirements.Core.Schema
{
    public class SchemaAdder
    {
        private const string SchemaDirectory = @"C:\Users\Amy\Documents\Projects\WordRequirements\src\Core\Schema";
        private const string FileName = "WordRequirements.xsd";
        private static readonly string SchemaPath = Path.Combine(SchemaDirectory, FileName);

        private const string Namespace = "http://ProgrammerGrrl.com/WordRequirements.xsd";
        private const string Alias = "Word Requirements";

        public void AddSchema()
        {
            var application = ApplicationContext.Application;
            var document = application.ActiveDocument;

            AddToApplication(application);
            AddToDocument(document);
        }

        private static void AddToApplication(Application application)
        {
            var namespaces = application.XMLNamespaces;

            if (!HasSchema(namespaces))
                namespaces.Add(SchemaPath, Namespace, Alias);
        }
        private static bool HasSchema(XMLNamespaces namespaces)
        {
            return namespaces
                .Cast<XMLNamespace>()
                .Any(ns => ns.URI == Namespace);
        }

        private static void AddToDocument(Document document)
        {
            var schemas = document.XMLSchemaReferences;

            if (!HasSchema(schemas))
                schemas.Add(Namespace);
        }

        private static bool HasSchema(XMLSchemaReferences schemas)
        {
            return schemas
                .Cast<XMLSchemaReference>()
                .Any(schema => schema.NamespaceURI == Namespace);
        }
    }
}
