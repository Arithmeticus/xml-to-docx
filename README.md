# xml-to-docx

## Why?

You're someone that handles XML (e.g., a publisher, a journal). 

You've got an XML file and you have questions about it. Or you want to discuss possible changes with other people.

The folks you need to work with do not have Oxygen or an XML editor. But they have Microsoft Word, OpenOffice, or another modern word processor.

You want to avoid at all costs a lossy conversion of your XML to a presentation format, have the discussion, then suffer another lossy conversion back to the original XML file. Just think of the headaches and the time lost repairing data lapses.

You don't want simply to paste the XML file as plain text in a Word document. That looks horrible. (Try it.) But you're getting closer to what you want.

Ideally, the Word document should have paragraph indentations that reflect the XML tree structure. Blocks of white space indentation should be dropped (perhaps). XML elements should be color coded, and perhaps reduced in size, to avoid confusion. Slighly reduce the opacity of the XML apparatus, so that elements, attributes, comments, and processing instructions can be seen, but do not steal the show. Or maybe you do want them to pop off the screen, and you want some control over that. 

That's what xml-to-docx is for. An XSLT application, xml-to-docx pushes any XML file into a Word docx file. Go ahead and track your changes, circulate comments, and do everything one does when doing traditional editing. Whenever you want to bring the file back to the original master XML format, simply copy the entire document and paste it as a plain text file. You might need to repair inadvertent errors that occurred in the editing process.

That's it. No lossy conversions (applied twice!). 

## Getting Started
1. **Make sure you can run an XSLT 3.0 application.**  xml-to-docx is an XSLT application. There's a built-in Oxygen project file, in case you prefer Oxygen. Or maybe you prefer the command line. Your choice. But it's XSLT 3.0. (Memo to developers: writing XSLT 1.0 code is like eating glass; please at least get up to 2.0.)
1. **Adjust parameters**. Adjust [`xml-to-docx.xsl`](xml-to-docx.xsl) for parameters you might want to change. There are loads of them, allowing you to change font size, color, transparency, or to play with space normalization. If you don't have experience editing an XSLT file, you might want to find someone who does, and work with them on your changes. There are 
1. **Run the transformation**. Pass your source XML file along with [`xml-to-docx.xsl`](xml-to-docx.xsl) into the XSLT processor of your choice. If you are working within Oxygen, open up the project. You'll see there's a transformation scenario, `run xml-to-docx`, that can be associated with whatever file you like.

## Workflow

1. **Run xml-to-docx**. Make sure it looks as you wish, then circulate the file to other people. 
1. **Doing your commenting/editing**. Simply treat the new file as your working document, making sure that changes keep the file well-formed. If you are working with people inexperienced with XML, warn them that they shouldn't change the colored parts, and they should keep track changes on.
1. **Bring the material back into XML**. The fastest way back is simply to select the entire text (ctrl+A), copy it (ctrl+C), then paste it in a plain text file. If there are any residual notes or comments, they will likely appear at the bottom of the file, and can be easily deleted, or incorporated into the XML structure. This is the fastest way back. You could also use your word processor to save as (export) a text file. But look out. Microsoft Word tends to save text files not as UTF-8 but as codepage 1252, and you'll likely get a lot of garbage in your output. The time it takes to configure and execute the export is much longer than the time it takes to run my suggested solution.

## Dependencies

xml-to-docx relies upon [Open and Save Archive](https://github.com/Arithmeticus/xslt-for-docx/), a simple library that allows you to open up any compressed archive, transform the contents, then perhaps save the archive again.  

In many projects, dependencies are handled with Maven, Gradle, or Git submodules. In the past those build solutions have given me headaches, and it's really overkill for a simple application like xml-to-docx. 

Instead, I have put the Open and Save Archive application directly in [`dependencies/open-and-save-docx.xsl`](dependencies/open-and-save-docx.xsl), along with a Schematron validator that will check to see if the application's top-level components are up to date. A Schematron quick fix will let you update components directly within Oxygen.

## Testing

Try opening the Oxygen project. Open `xml-to-docx.xsl`. Hit ctrl+shft+T. If the settings are set to default, you'll get a new file, `xml-to-docx.xsl.docx`. Open it up in the word processor of choice.

## About

Version: 2021-01

license: [GNU General Public License](https://opensource.org/licenses/GPL-3.0)

author: Joel Kalvesmaki

