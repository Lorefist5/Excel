using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;
using System.Text;
using System.Linq;

[Generator]
public class ExcelPropertyGenerator : ISourceGenerator
{
    public void Initialize(GeneratorInitializationContext context)
    {
        // Optional: Initialize logging or setup required by the generator
    }

    public void Execute(GeneratorExecutionContext context)
    {
        // Iterate over all syntax trees in the compilation
        foreach (var syntaxTree in context.Compilation.SyntaxTrees)
        {
            var semanticModel = context.Compilation.GetSemanticModel(syntaxTree);
            var classDeclarations = syntaxTree.GetRoot().DescendantNodes().OfType<ClassDeclarationSyntax>();

            foreach (var classDecl in classDeclarations)
            {
                var classSymbol = semanticModel.GetDeclaredSymbol(classDecl) as INamedTypeSymbol;
                if (classSymbol == null) continue;

                foreach (var propertySymbol in classSymbol.GetMembers().OfType<IPropertySymbol>())
                {
                    var excelAttribute = propertySymbol.GetAttributes().FirstOrDefault(attr => attr.AttributeClass?.ToDisplayString() == "YourNamespace.ExcelAttribute");
                    if (excelAttribute != null)
                    {
                        var typeArgument = excelAttribute.NamedArguments.FirstOrDefault(arg => arg.Key == "Type").Value;
                        if (typeArgument.Value is ITypeSymbol typeSymbol && propertySymbol.Type.Name != typeSymbol.Name)
                        {
                            var propertyType = typeSymbol.ToDisplayString();
                            var propertyName = propertySymbol.Name + "Int";  // Adjust name based on your convention
                            var sourceCode = GenerateProperty(classSymbol.Name, propertyName, propertySymbol.Name, propertyType);

                            context.AddSource($"{classSymbol.Name}_{propertySymbol.Name}_Extension.cs", SourceText.From(sourceCode, Encoding.UTF8));
                        }
                    }
                }
            }
        }
    }

    private string GenerateProperty(string className, string propertyName, string originalPropertyName, string propertyType)
    {
        // Adjusted to properly check conversion logic and use it in generated code
        return $@"
public partial class {className}
{{
    [Excel(IsProperty = false)]
    public {propertyType} {propertyName} => {propertyType}.TryParse(this.{originalPropertyName}, out var temp) ? temp : default({propertyType});
}}
";
    }
}
