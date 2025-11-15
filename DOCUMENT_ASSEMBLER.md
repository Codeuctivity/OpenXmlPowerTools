# OpenXmlPowerTools

OpenXmlPowerTools provides guidance and example code for programming with Open XML Documents (DOCX, XLSX, and PPTX). It is based on, and extends the functionality of the Open XML SDK.

## Document Assembler

The `DocumentAssembler` module is a powerful tool for generating documents from templates and data. It allows you to create `.docx` files with dynamic content, such as tables, conditional sections, and repeating elements, based on data from an XML file.

### Key Features

*   **Template-Based Document Generation**: Create documents from Word templates (`.docx`) and populate them with data from XML files.
*   **Content Replacement**: Use simple placeholders in your template to insert data from your XML file.
*   **Dynamic Tables**: Automatically generate tables in your document based on data from your XML file.
*   **Conditional Content**: Include or exclude parts of your document based on conditions in your data.
*   **Repeating Content**: Repeat sections of your document for each item in a collection in your data.
*   **Error Handling**: The `DocumentAssembler` will report errors in the generated document if it encounters any issues with your template or data.

### How it Works

The `DocumentAssembler` works by processing a Word document that contains special markup in content controls or in paragraphs. This markup defines how the document should be assembled based on the provided XML data.

The process is as follows:

1.  **Create a Template**: Start with a regular Word document (`.docx`).
2.  **Add Placeholders**: Use content controls or special syntax in paragraphs to define placeholders for your data.
3.  **Provide Data**: Create an XML file that contains the data you want to insert into the document.
4.  **Assemble the Document**: Use the `DocumentAssembler.AssembleDocument` method to merge the template and data, producing a new Word document.

### Template Syntax

The template syntax uses XML elements within content controls or as text in the format `<#ElementName ... #>`.

#### Content Replacement

To replace a placeholder with a value from your XML data, you can use the `Content` element. The `Select` attribute contains an XPath expression to select the data from the XML file.

**Example:**

`<#Content Select="Customer/Name" #>`

#### Tables

To generate a table, you use the `Table` element. The `Select` attribute specifies the collection of data to iterate over. The table in the template must have a prototype row, which will be repeated for each item in the data.

**Example:**

`<#Table Select="Customers/Customer" #>`

#### Conditional Content

You can conditionally include content using the `Conditional` element. The `Select` attribute specifies the data to test, and the `Match` or `NotMatch` attribute specifies the value to compare against.

**Example:**

`<#Conditional Select="Customer/Country" Match="USA" #>
... content to include if the customer is from the USA ...
<#EndConditional #>`

#### Repeating Content

To repeat a section of the document, you can use the `Repeat` element. The `Select` attribute specifies the collection of data to iterate over.

**Example:**

`<#Repeat Select="Customers/Customer" #>
... content to repeat for each customer ...
<#EndRepeat #>`

### Getting Started

To use the `DocumentAssembler`, you will need to:

1.  Add a reference to the `OpenXmlPowerTools` library in your project.
2.  Create a Word template with the appropriate placeholders.
3.  Create an XML data file.
4.  Call the `DocumentAssembler.AssembleDocument` method to generate your document.

For more detailed examples and documentation, please refer to the `DocumentAssembler`, `DocumentAssembler01`, `DocumentAssembler02`, and `DocumentAssembler03` projects in the `OpenXmlPowerToolsExamples` directory.