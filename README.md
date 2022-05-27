# awesome-vba

## A note on symbology

To help you in finding projects suitable for you, this awesome list uses the following symbology

#### Platform Compatibility

[p_all]: # "Compatible on all platforms"
[p_mac]: # "Mac OS only"
[p_win]: # "Windows OS only"

* [👑][p_all] - Compatible on all platforms
* [🍎][p_mac] - Mac compatible
* [🖼][p_win] - Windows compatible

#### Application compatibility 

[a_all]: #  "All applications"
[a_wd]: #   "Word"
[a_xl]: #   "Excel"
[a_ac]: #   "Access"
[a_ol]: #   "Outlook"
[a_pp]: #   "PowerPoint"


* [⭐][a_all] - All applications
* [✒️][a_wd] - Word
* [📊][a_xl] - Excel
* [🅰️][a_ac] - Access
* [📧][a_ol] - Outlook
* [🎞️][a_pp] - Powerpoint
* [🦆](# "Misc") - Miscellaneous applications (MS Project, AutoCAD, etc.) - Specify in short description

#### Other important information

[o_32]:   #  "32-bit only"
[o_pass]: #  "VBA is password protected"  

* [🏺][o_32] - 32-bit only 
* [🔒][o_pass] - Written in VBA but the code is password protected
* [👽](# "Requires external dependencies") - Requires external dependencies e.g. DLLs
* [💣](# "Requires installation") - Requires installation
* [💲](# "Link includes/leads to paid content") - Link includes/leads to paid content

## Contents

- [awesome-vba](#awesome-vba)
  - [A note on symbology](#a-note-on-symbology)
      - [Platform Compatibility](#platform-compatibility)
      - [Application compatibility](#application-compatibility)
      - [Other important information](#other-important-information)
  - [Contents](#contents)
  - [Frameworks](#frameworks)
  - [Libraries](#libraries)
    - [Data Formats](#data-formats)
      - [JSON](#json)
      - [CSV](#csv)
      - [XML](#xml)
    - [Data Structures](#data-structures)
      - [Array-List](#array-list)
      - [Dictionary](#dictionary)
    - [Math libraries](#math-libraries)
    - [Database tools](#database-tools)
    - [Userform tools](#userform-tools)
    - [Memory tools](#memory-tools)
    - [Web tools](#web-tools)
  - [Developer tools](#developer-tools)
  - [Examples](#examples)
    - [UI Ribbon](#ui-ribbon)
    - [UI Userforms](#ui-userforms)
    - [VBE UI](#vbe-ui)
    - [Games](#games)
  - [Style Guides](#style-guides)
  - [Resources](#resources)
    - [Books / Websites](#books--websites)
    - [Youtube](#youtube)
    - [Forums](#forums)
  - [Contributing](#contributing)

------

## Frameworks

* [🖼][p_win][⭐][a_all] [stdVBA](http://github.com/sancarn/stdVBA) - A framework containing numerous classes for automation and utility. Focuses on code compactness and long-term maintainability.
* [🖼][p_win][⭐][a_all][🏺][o_32] [VbCorLib](https://github.com/kellyethridge/VBCorLib) - A framework which brings many powerful .NET classes to VBA/VB6.

## Libraries

### Data Formats

#### JSON

* [👑][p_all][⭐][a_all] [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) - JSON conversion and parsing.

#### CSV

* [👑][p_all][⭐][a_all] [VBA-CSV-interface](https://github.com/ws-garcia/VBA-CSV-interface) - Powerful, fast and comprehensive RFC-4180 compliant CSV/TSV/DSV data management library.

#### XML

* [👑][p_all][⭐][a_all] [VBA-XML](https://github.com/VBA-tools/VBA-XML) - XML conversion and parsing.

### Data Structures

#### Array-List

* [👑][p_all][⭐][a_all] [Better array](https://github.com/Senipah/VBA-Better-Array/tree/master/src) - An array class providing features found in more modern languages


#### Dictionary

* [👑][p_all][⭐][a_all] [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary) - A dictionary object which stores key-value pairs.
* [🖼][p_win][⭐][a_all] [VBA-ExtendedDictionary](https://github.com/SSlinky/VBA-ExtendedDictionary) - A dictionary object using Scripting.Dictionary but exposes some additional useful functionality.

### Math libraries

* [👑][p_all][⭐][a_all] [VBA-Math-Objects](https://github.com/Beakerboy/VBA-Math-Objects) - A matrix and vector library.

### Database tools

* [SQL Library](https://github.com/Beakerboy/VBA-SQL-Library) - An OOP SQL Library for psql, mssql, mysql databases.

### Userform tools

* [🖼][p_win][⭐][a_all] [Material UI](https://github.com/todar/VBA-Material-Design) - Make your userform feel modern with Material UI.
* [👑][p_all][⭐][a_all] [Easy EventListener](https://github.com/todar/VBA-Userform-EventListener) - Consolodate all event handling of a userform into 1 callback.
* [🖼][p_win][⭐][a_all][👽](# "Requires external DLL") [Modern UI Components](https://github.com/krishKM/Modern-UI-Components-for-VBA) - Custom modern looking controls. 
* [🖼][p_win][⭐][a_all] [MVVM](https://github.com/rubberduck-vba/MVVM) - Model-View-ViewModel Infrastructure for maintainable userform development.
* [🖼][p_win][⭐][a_all] [VBA Userform Transitions and Animations](https://github.com/todar/VBA-Userform-Animations) - An excellent library for implementing animation easings into the Userform.

### Memory tools

* [👑][p_all][⭐][a_all] [VBA-MemoryTools](https://github.com/cristianbuse/VBA-MemoryTools)

### Web tools

* [👑][p_all][⭐][a_all] [VBA-Web](https://github.com/VBA-tools/VBA-Web) - Connect VBA, Excel, Access, and Office for Windows and Mac to web services and the web

## Developer tools

* [🖼][p_win][⭐][a_all][💣](# "Requires installation") [Rubberduck](https://rubberduckvba.com/) - An open-source COM add-in project that integrates with the Visual Basic Editor to add modern-day features to the familiar IDE. Works in VBA6, VBA7.x (x86/x64), and yes, in VB6 too!
* [👑][p_all][⭐][a_all] [VBA-IDE-Code-Export](https://github.com/spences10/VBA-IDE-Code-Export) - Addin contains a code importer and exporter for use with git (or any VCS)

## Examples

### UI Ribbon

* [Ron de Bruin - Ribbons/QAT](https://www.rondebruin.nl/win/section2.htm) - A leading resource for information/samples on developing custom ribbons.

### UI Userforms

TBC

### VBE UI

### AddIns

* [MenuRighter](https://yoursumbuddy.com/blog/menurighter/)
* [Custom UI XML Editor](https://yoursumbuddy.com/ribbon-customui-xml-editor/) - Addin for directly adding, editing and validating ribbon XML (Excel 2010+).

TBC

### Games

TBC

## Style Guides

* [VBA Standard](https://sslinky.github.io/VBA-Standard/) - Guide prepared by moderators of the r/vba subreddit.
* [todar's style guide](https://github.com/todar/VBA-Style-Guide) 
* [RubberDuck's style guide](https://rubberduckvba.wordpress.com/2021/05/29/rubberduck-style-guide/)

## Resources

### Books / Websites

* [MSDN VBA Documentation](https://msdn.microsoft.com/en-us/vba/office-vba-reference)
* [MS-VBAL Language Spec](https://docs.microsoft.com/en-gb/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74)
* [VB6 Coding conventions](https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa240822(v%3dvs.60))
* [Ron de Bruin](https://www.rondebruin.nl/index.htm) - Simple-Intermediate topics.
* [Bytecomb VBA Reference](https://bytecomb.com/vba-reference/) - Intermediate-advanced topics.
* [Chip Pearson's website](http://www.cpearson.com/excel) - Great resource for beginners-intermediate.
* [VBA for smarties](http://www.snb-vba.eu/inhoud_en.html) - A great reference to numerous data structures and mechanics.
* [💲](# "Some cheatsheets are paid-for content")[Automate Excel's cheat sheets](https://www.automateexcel.com/vba/cheatsheets/)
* [💲](# "Have to pay for book")[Excel VBA Programming For Dummies book](https://www.google.com/search?q=Excel+VBA+Programming+For+Dummies+book)
* [VBA E-Book for beginners](https://goalkicker.com/VBABook/)
* [Excel VBA E-Book for beginners](https://goalkicker.com/ExcelVBABook/)
* [Rubberduck Blog](https://rubberduckvba.wordpress.com/) - Intermediate-Advanced topics.

### Youtube

* [Excel Macro Mastery](https://www.youtube.com/c/Excelmacromastery) - Youtube channel for Paul Kelly (MVP) - excelmacromastery.com. 
* [Sigma Coding](https://www.youtube.com/c/SigmaCoding) - Large catalogue of tutorials - beginner through to advanced. Delves into interesting areas of VBA and it uses not explored by other content creators.
* [WiseOwl's VBA tutorials](https://www.youtube.com/playlist?list=PLNIs-AWhQzckr8Dgmgb3akx_gFMnpxTN5) - Great all-round resource for VBA. Perfect introduction for beginners. In-depth lessons into all aspects of VBA. Huge playlist that covers most types of VBA. 
* [💲](# "Some libraries used are non-FOSS and created by VBA A2Z") [VBA A2Z](https://www.youtube.com/c/VBAA2Z) - Many tutorials, some paid content. Good array of intereting and different topics - in-depth tutorials into different parts of VBA, with some .NET/VSTO videos. Strong focus on UI development.
* [Excel VBA Is Fun](https://www.youtube.com/c/ExcelVbaIsFun)
* [Excel for Freelancers](https://www.youtube.com/c/ExcelForFreelancers) - Hands-on tutorials to developing specific applications from beginning through to end. All levels.
* [Leila Gharani](https://www.youtube.com/c/LeilaGharani) - Office-wide focus - useful for beginners.


### Forums

* [Reddit](http://reddit.co.uk/r/vba) - Daily VBA Q&A. Occasional Pro-Tip sharing and Show & Tell library publishing.
* [Visual Basic Discord](https://discord.gg/gpcSue9f) - A chat room for VB.NET/VBA/VB6 fanatics.
* [Excel Discord](https://discord.gg/PU2vVDeb) - Focus on helping Excel users.
* [MrExcel](https://www.mrexcel.com/board/) - Mostly Excel generic, but a lot of VBA content can be found here also.
* [Excel Forum](https://www.excelforum.com/)
* [📧][a_ol] [Slipstick](https://www.slipstick.com/outlook) - Excellent forum for the Slipstick website (Outlook VBA) of Diane Poremsky (MVP). Diane is quick to respond, and her answers are extremely helpful.
* [VBForums - Office Development](https://www.vbforums.com/forumdisplay.php?37-Office-Development) - Forum with focus on VB6/.NET with VBA section.

## Contributing

Your contributions are always welcome! Please take a look at the [contribution guidelines](./Contributing.md) first.

