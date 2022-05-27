# Contributing

I will keep some pull requests open if I'm not sure whether those libraries are awesome, you could vote for them by adding 👍 to them. 
Pull requests will be merged when their votes reach 5.

* Add one link per Pull Request.
* Make sure the PR title is in the format of `Add project-name`.
* Write down the reason why the library is awesome.
* Add the link: `* <symbology> [project-name](http://example.com/) - A short description ends with a period.`
  * Replace `<symbology>` with symbols specifying application and OS compatibility. [See below](#Symbology)
* Keep descriptions concise and short.
* Add a section if needed.
* Add the section description.
* Add the section title to Table of Contents.
* Search previous Pull Requests or Issues before making a new one, as yours may be a duplicate.
* Don't mention VBA in the description as it's implied.
* Check your spelling and grammar.
* Remove any trailing whitespace.

### Symbology

This repository uses symbology to indicate restrictions in compatibility. This is to help users find projects which work for their particular use case. Symbology should match the reality of a project rather than it's ambition. For instance `stdVBA` aspires to be multi-platform, but it currently lacks a lot of Mac compatibility. In this scenario it's labelled as 🖼.

#### 1. Specifying Platform Compatibility

Platform compatibility is especially important for Mac users as many libraries are windows only. If you've used `CreateObject` to create an object e.g. `Scripting.Dictionary`, `VBScript.Dictionary` etc. then your library is likely Windows OS Only. Additionally if you've used external DLL functions, the likelihood is your library is Windows only. In order to make these mac-compatible the library needs to use Mac-native functions from libc or objc.

* 👑 - Compatible on all platforms
* 🍎 - Mac OS only
* 🖼 - Windows OS only

#### 2. Specifying Application compatibility 

If a library is built for and/or only works within a specific application and/or relies on the application running specify as below. If no 

* ⭐ - All applications
* ✒️ - Word
* 📊 - Excel
* 🅰️ - Access
* 📧 - Outlook
* 🎞️ - Powerpoint
* 🦆 - Miscellaneous applications (MS Project, AutoCAD, etc.) - Specify in short description

#### 3. Specifying other important information

Many people use VBA in business environments because they don't have better tools available. Dependency download may be blocked, or installation may be something that can only be done by IT staff.

* 👽 - Requires external dependencies e.g. DLLs
* 💣 - Requires installation
* 🏺 - 32-bit/VB6 only 
* 💲 - Link includes/leads to paid content
* 🔒 - VBA source code is password protected and/or hidden.

#### 4. Symbology should contain tooltips

As suggested in #1 tooltips should be added to symbology to further help new users browser the awesome list.

```md
* [👑](# "Compatible on all platforms")[⭐]("All applications")
```

In order to keep the document clean, several IDs have been added for common tooltips:

```
[p_all]:  #  "Compatible on all platforms"
[p_mac]:  #  "Mac OS only"
[p_win]:  #  "Windows OS only"
 
[a_all]:  #  "All applications"
[a_wd]:   #  "Word"
[a_xl]:   #  "Excel"
[a_ac]:   #  "Access"
[a_ol]:   #  "Outlook"
[a_pp]:   #  "PowerPoint"
 
[o_32]:   #  "32-bit only"
[o_pass]: #  "VBA is password protected"  
```

These can be used as follows:

```md
* [👑][p_all][⭐][a_all]
```

Tooltips can be modified to give further helpful detail and should be considered especially for 👽, 💲 and 🦆.

```md
* [👑][p_all][⭐][a_all][👽](# "Requires external DLL")
* [👑][p_all][⭐][a_all][💲](# "Some non-FOSS cheatsheets")
* [👑][p_all][⭐][a_all][🦆](# "Works in Auto-CAD")
```



### Symbology Examples

* 👑⭐ - Compatible on all operating systems and in all applications
* 🖼✒️📊 - Only compatible on windows and only works in Word and Excel.

### Contribution examples

```
* [🖼](# "Windows OS only")[⭐](# "All applications") [stdVBA](http://github.com/sancarn/stdVBA) - A framework containing numerous classes for automation and utility. Focuses on code compactness and long-term maintainability.
```