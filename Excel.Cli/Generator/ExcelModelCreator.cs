using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace Excel.Cli.Generator;

public class ExcelModelCreator
{

    public string CreateModel(string excelFilePath, string className, string sheetName = "Sheet1", int headerRow = 1, int headerColumn = 1)
    {
        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            var worksheet = package.Workbook.Worksheets[sheetName]; // use the specified worksheet
            return GenerateClassFromWorksheet(worksheet, className, headerRow, headerColumn);
        }
    }

    public Dictionary<string, string> CreateModelsForAllSheets(string excelFilePath, int headerRow = 1, int headerColumn = 1)
    {
        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            var models = new Dictionary<string, string>();

            foreach (var worksheet in package.Workbook.Worksheets)
            {
                var className = ToPascalCase(worksheet.Name);
                var model = GenerateClassFromWorksheet(worksheet, className, headerRow, headerColumn);
                models[worksheet.Name] = model;
            }

            return models;
        }
    }
    public void GenerateProperties(string classFilePath)
    {
        var tree = CSharpSyntaxTree.ParseText(File.ReadAllText(classFilePath));
        var root = tree.GetRoot();

        var classDeclaration = root.DescendantNodes().OfType<ClassDeclarationSyntax>().First();
        var newMembers = new List<MemberDeclarationSyntax>();

        foreach (var property in classDeclaration.Members.OfType<PropertyDeclarationSyntax>())
        {
            var excelAttribute = property.AttributeLists.SelectMany(al => al.Attributes)
                .FirstOrDefault(a => a.Name.ToString() == "Excel");

            if (excelAttribute != null)
            {
                var isPropertyArgument = excelAttribute.ArgumentList.Arguments
                    .FirstOrDefault(a => a.NameEquals.Name.Identifier.Text == "IsProperty");

                if (isPropertyArgument != null && ((LiteralExpressionSyntax)isPropertyArgument.Expression).Token.ValueText == "false")
                {
                    // Skip this property if it has [Excel(IsProperty = false)]
                    continue;
                }

                var typeArgument = excelAttribute.ArgumentList.Arguments
                    .FirstOrDefault(a => a.NameEquals.Name.Identifier.Text == "Type");

                if (typeArgument != null)
                {
                    var typeName = ((TypeOfExpressionSyntax)typeArgument.Expression).Type.ToString();
                    var propertyName = property.Identifier.Text + typeName;

                    var parseInvocation = SyntaxFactory.ParseExpression($"{typeName}.Parse({property.Identifier.Text})");
                    var arrowClause = SyntaxFactory.ArrowExpressionClause(parseInvocation);

                    var newProperty = SyntaxFactory.PropertyDeclaration(SyntaxFactory.ParseTypeName(typeName), propertyName)
                        .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword))
                        .WithExpressionBody(arrowClause)
                        .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken));

                    var attribute = SyntaxFactory.Attribute(SyntaxFactory.IdentifierName("Excel"), SyntaxFactory.AttributeArgumentList(SyntaxFactory.SingletonSeparatedList(SyntaxFactory.AttributeArgument(SyntaxFactory.NameEquals(SyntaxFactory.IdentifierName("IsProperty")), null, SyntaxFactory.LiteralExpression(SyntaxKind.FalseLiteralExpression)))));
                    newProperty = newProperty.AddAttributeLists(SyntaxFactory.AttributeList(SyntaxFactory.SingletonSeparatedList(attribute)));

                    newMembers.Add(newProperty);
                }
            }
        }

        var newClassDeclaration = classDeclaration.AddMembers(newMembers.ToArray());
        var newRoot = root.ReplaceNode(classDeclaration, newClassDeclaration);

        var newClassCode = newRoot.NormalizeWhitespace().ToFullString();
        File.WriteAllText(classFilePath, newClassCode);
    }

    private string GenerateClassFromWorksheet(ExcelWorksheet worksheet, string className, int headerRow, int headerColumn)
    {
        var classDeclaration = SyntaxFactory.ClassDeclaration(className)
            .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword));

        for (int i = headerColumn; i <= worksheet.Dimension.End.Column; i++)
        {
            if (worksheet.Cells[headerRow, i].Value == null)
            {
                continue;
            }
            var headerName = worksheet.Cells[headerRow, i].Value.ToString();
            var propertyName = ToPascalCase(headerName); // Convert to PascalCase

            // Special case: if the header name starts with a number followed by '+', move the number to the end
            var match = Regex.Match(propertyName, @"^(\d+)\+(.*)$");
            if (match.Success)
            {
                propertyName = match.Groups[2].Value + match.Groups[1].Value;
            }

            // Remove special characters
            propertyName = Regex.Replace(propertyName, @"[^a-zA-Z0-9_]", "");

            // Skip if the property name is empty or starts with a number
            if (string.IsNullOrEmpty(propertyName) || char.IsDigit(propertyName[0]))
            {
                continue;
            }

            var propertyDeclaration = SyntaxFactory.PropertyDeclaration(SyntaxFactory.ParseTypeName("string"), propertyName)
                .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword))
                .AddAccessorListAccessors(
                    SyntaxFactory.AccessorDeclaration(SyntaxKind.GetAccessorDeclaration).WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken)),
                    SyntaxFactory.AccessorDeclaration(SyntaxKind.SetAccessorDeclaration).WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken)));

            // Add Excel attribute with Name property
            var nameEquals = SyntaxFactory.NameEquals(SyntaxFactory.IdentifierName("Name"));
            var attributeArgument = SyntaxFactory.AttributeArgument(nameEquals, null, SyntaxFactory.LiteralExpression(SyntaxKind.StringLiteralExpression, SyntaxFactory.Literal(headerName)));
            var attributeArgs = SyntaxFactory.AttributeArgumentList(SyntaxFactory.SingletonSeparatedList(attributeArgument));

            var attribute = SyntaxFactory.Attribute(SyntaxFactory.IdentifierName("Excel"), attributeArgs);
            propertyDeclaration = propertyDeclaration.AddAttributeLists(SyntaxFactory.AttributeList(SyntaxFactory.SingletonSeparatedList(attribute)));

            classDeclaration = classDeclaration.AddMembers(propertyDeclaration);
        }

        var namespaceDeclaration = SyntaxFactory.NamespaceDeclaration(SyntaxFactory.ParseName("YourNamespace"))
            .AddMembers(classDeclaration);

        var code = namespaceDeclaration
            .NormalizeWhitespace() // to format the code properly
            .ToFullString();

        return code;
    }
    private string ToPascalCase(string input)
    {
        // Replace spaces with underscores, then use a regex to convert to PascalCase
        string withUnderscores = input.Replace(" ", "_");
        return Regex.Replace(withUnderscores, "(?:^|_)(.)", match => match.Groups[1].Value.ToUpper());
    }
}
