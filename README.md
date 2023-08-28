# Metadata Export-Import

This is a JavaScript extension for [Adobe Bridge](https://www.adobe.com/products/bridge.html) 2015 and higher exports and imports [XMP metadata](https://www.adobe.com/products/xmp.html) embedded in media files using a tab-delimited text file (spreadsheet).

Media files and metadata are matched by file name and file path.

In the script window, click the "?" button for instructions

Default metadata field sets:
 - Basic (Title, Creator, Description, Keywords, Lightroom-Keywords, Copyright Notice, Copyright Status, File-Date Created, File-Date Modified, File-Date Original)
 - [IPTC Core](https://www.iptc.org/std/photometadata/specification/IPTC-PhotoMetadata#iptc-core-schema-1-4-specifications)
 - [Dublin Core](https://www.dublincore.org/specifications/dublin-core/dcmi-terms/#section-3)
 - [VRA Core (artworks/visual culture)](https://docs.google.com/spreadsheets/d/1f4UU-4tnlq_DlqgcQ2xYHsKfvp3tP-82Re5hr0WsTZw/edit?usp=sharing)
 
This list can be edited and custom field sets can be added. See **Customize Metadata Fields** below for more details.

Supported file types: ai, avi, bmp, dng, flv, gif, indd, indt, jpg, mp2, mp3, mp4, mov, pdf, png, psd, swf, tif, wav, wma, wmv, xmp, and RAW files using XMP sidecar

This extension is written in [Adobe ScriptUI](https://estk.aenhancers.com/index.html) and utilizes the [Adobe Bridge XMP scripting API](https://estk.aenhancers.com/10%20-%20Scripting%20Access%20to%20XMP%20Metadata/accessing-the-xmp-scripting-api.html)

## Installation
 1. Start Adobe Bridge
 2. Go to Bridge Preferences
 
    Windows
      - Go to Edit --> Preferences --> Startup Scripts
    
    Mac:
      - Go to Adobe Bridge --> Preferences --> Startup Scripts
	  
 3. Click "Reveal My Startup Scripts".  This will open the correct folder
 4. In this folder, copy the the downloaded file "metadata_export_import.jsx"
 5. Close the folder
 6. Quit Bridge
 7. Start Bridge
 8. When prompted, confirm the installation of "metadata_export_import"

## Open the plugin in Adobe Bridge
  1. Locate the "Metadata Deluxe" menu at the top of the Bridge window
  2. Click "Export-Import" on the dropdown
  3. Select which Metadata Fields to use
  4. Select Export or Import
  5. In the plugin window, click the "?" button for instructions

### Export

- Choose which files to export
 - Selected files only
    - Select file(s) in Adobe Bridge
    - Under ‘Images’, choose 'Selected thumbnails in Bridge'
    - Select desired export options (see below)'
    - click 'Export'
    - Choose name and location for your export file (.txt format)
    - Save

 -  All image files in a folder:
    - Select 'Entire folder'
    - Use 'browse'  to select folder you wish to use
    - If you want to export images in all subfolders check the box 'include subfolders'
    - Select desired export options (see below)
    - Click 'Export'
    - Choose name and location for your export file (.txt format)
    - Save

- After export is complete:
 - Open your .txt file in a spreadsheet using tab delimiters
 - Because the .txt file is created as UTF-8, it is best to use the Excel import text wizard to retain
proper character encoding, e.g. ©, 작, ü
   - Use import option and select a .txt file to import
   - Text Import Wizard
     - File origin = 65001- Unicode (UTF-8) or (UTF-16)
     - Delimited = Tab
     - Finish
  - The first row will contain the field names
  - Each row below that will contain metadata for a single exported file
		 
### Import

This script imports data from a tab delimited text file into image files with matching filenames

If you have already exported metadata and edited it in Excel and want to import it back in:
 - Select the folder containing the target images
   - If you have nested folders, check 'include subfolders'
 - Select the text file to be imported
   - Click 'Import'
   
If you are starting with external metadata (from a database, etc.):
 - First, you need a properly formatted spreadsheet.
   - Open the Export window and export the files you want to work on
   - Import this text file into a spreadsheet
   - This will give you the file names and columns to enter your metadata into		 
   - Look for this new file on your desktop
     - Open the text file in Excel
     - Add data to the appropriate columns
     - Save as a tab-delimited text (.txt) file
      - Format as Unicode UTF-8 or 16
      
To import data into image files:
 - Select the folder containing the target images
   - If you have nested folders, check 'include subfolders'
 - Select the text file to be imported
 - Click 'Import' and cross your fingers
 
## Customize Metadata Fields

The set list of metadata fields to be exported or imported can be customized by editing the script in a JavaScript code editor (a plain text editor will work too).

Each list of fields, e.g., IPTC Core, is an array of property objects. Each of these arrays (lists) is added to the Metadata Fields dropdown list (fieldsSelectDrp) for selection in the user interface.

Included field arrays:
 - basicFieldsArr (commonly used fields)
 - iptcFieldsArr (IPTC Core)
 - dcFieldsArr (Dublin Core)
 - vraFieldsArr (VRA Core - artworks/visual culture)

You can create a custom field array by combining properties from the included arrays or defining new properties.

Required property object format (this defines a single XMP property):
```
{Schema_Field:' ', Label:' ', Namespace:' ', XMP_Property:' ', XMP_Type:' '}
```
  
Example custom list array:
```
[
 {Schema_Field:'File-Name', Label:'File Name', Namespace:'file', XMP_Property:'file', XMP_Type:'file'},
 {Schema_Field:'File-Folder', Label:'File Folder', Namespace:'file', XMP_Property:'file', XMP_Type:'file'}, 
 {Schema_Field:'IPTC-Creator', Label:'Creator', Namespace:'http://purl.org/dc/elements/1.1/', XMP_Property:'dc:creator', XMP_Type:'seq'},
 {Schema_Field:'DC-Publisher', Label:'Publisher', Namespace:'http://purl.org/dc/elements/1.1/', XMP_Property:'dc:publisher', XMP_Type:'bag'},
 {Schema_Field:'VRA-Work Agent', Label:'Work Agent', Namespace:'http://www.vraweb.org/vracore/4.0/essential/', XMP_Property:'vrae:work.agent', XMP_Type:'text'},
 {Schema_Field:'Custom', Label:'Classification, Namespace:'http://www.myschema.net/', XMP_Property:'my:class', XMP_Type:'text'}
]
```

Required property object definitions:
 - Schema_Field:
   - The schema and field name (supplied by this tool) or 'Custom'
 - Label
   - Your label for the field (call the field anything you like)
 - Namespace
   - URL representing the schema (see schema specification or define your own)
   - Must end in '/' or '#'. Does not have to resolve to an actual web page
 - XMP_Property
   - 'prefix:property name' (see schema specification or define your own for Custom)
   - Must have a prefix, a colon, a property name. Must no have spaces. Example: 'dc:title'
   - prefix should be the preferred or commonly used prefix for a namespace (see schema specification)
   - property name must exactly match the schema specification (case-sensitive)
 - XMP_Type
   - text = Plain text string (e.g., single character, word, phrase, sentence, or paragraph)
   - bag = Unordered array of text strings (e.g., keywords in no particular order)
   - seq = Ordered array of text strings (e.g., list of authors in order of importance)
   - langAlt = Array of text strings in different languages using xml:lang qualifiers such as 'en-us'. 'x-default' is used when the language is undefined.
   - boolean = 'True' or 'False'
   - date = YYYY-MM-DD, YYYY-MM, YYYY

For more information on XMP properties, see the XMP Specification, https://www.adobe.com/devnet/xmp.html

To customize the fields dropdown list (fieldsSelectDrp):
 - If you want to add custom lists of fields, create those in a new property object (see above definition)
 - Edit var fieldsList to contain the field lists you want to use
 - Search for CUSTOMIZE HERE for elements that need to be changed

