# Contributing

Thanks for helping make Awesome‑VBA truly awesome!  Whether you write VBA, VB6 or both, we welcome your pull‑requests.

## Quick checklist

1. Try to stick to one link per PR. Small, focused PRs are easier to review.
2. PR title: `Add <project‑name>`.
3. Entry format (single line, bullet list):
   `- <symbology> [project‑name](https://…) – Short description ending with a period.`
4. Keep the description concise and short (≤ 120 chars is a good rule of thumb)
5. Add a new section & Table of Contents entry if none fits.
6. Search existing Issues/PRs first to avoid duplicates.
7. Proof‑read for spelling/grammar and remove trailing whitespace.

Pull‑requests will be generally accepted by maintainers. In some cases we will await 👍 reactions from the community/maintainers/contributors. Feel free to add your vote to open PRs and issues.

### Symbology

This awesome list uses icons to flag platform compatibility, host application, and other constraints so readers can quickly determine whether the project suits their environment and needs. Symbology should match the reality of a project rather than it's ambition. For instance `stdVBA` aspires to be multi-platform, but it currently lacks a lot of Mac compatibility. In this scenario it's labelled as [![p_win]](#-) [![p_nom]](#-).

Symbology should be of the form:

```md
- <platform-compatiblity> <application-compatibility> <other-constraints> [<Title>](...) - <Description>
```

#### 1. Platform compatibility

Platform compatibility is especially important for Mac users as many libraries are windows only. If you've used `CreateObject` to create an object e.g. `Scripting.Dictionary`, `VBScript.Dictionary` etc. then your library is likely Windows OS Only. Additionally if you've used external DLL functions, the likelihood is your library is Windows only. In order to make these mac-compatible the library needs to use Mac-native functions from libc or objc.

| Icon set                      | Mark-up                         | Description                        |
|-------------------------------|---------------------------------|------------------------------------|
| [![p_win]](#-) [![p_mac]](#-) | `[![p_win]](#-) [![p_mac]](#-)` | Compatible on both Windows and Mac |
| [![p_win]](#-) [![p_nom]](#-) | `[![p_win]](#-) [![p_nom]](#-)` | Compatible on Windows only         |
| [![p_now]](#-) [![p_mac]](#-) | `[![p_now]](#-) [![p_mac]](#-)` | Compatible on Mac only             |

#### 2. Host compatibility 

If a library is built for and/or only works within a specific application and/or relies on the application running specify as below.

| Icon                          | Mark-up                         | Description                        |
|-------------------------------|---------------------------------|------------------------------------|
| [![a_all]](#-)                | `[![a_all]](#-)`                | All applications                   |
| [![a_wd]](#-)                 | `[![a_wd]](#-)`                 | Word                               |
| [![a_xl]](#-)                 | `[![a_xl]](#-)`                 | Excel                              |
| [![a_ac]](#-)                 | `[![a_ac]](#-)`                 | Access                             |
| [![a_ol]](#-)                 | `[![a_ol]](#-)`                 | Outlook                            |
| [![a_pp]](#-)                 | `[![a_pp]](#-)`                 | PowerPoint                         |
| [![a_misc]](#- 'Misc')        | `[![a_misc]](#- 'Misc')`        | Miscellaneous applications (MS Project, AutoCAD, VB6, Python etc.) - Specify in short description |

#### 3. Other flags

Many people use VBA in business environments because they don't have better tools available. Dependency download may be blocked, or installation may be something that can only be done by IT staff. Additionally, sometimes libraries cost money, and thus require a cost center, preventing buy-in. This symbology aids users in understanding this.

| Icon                                                  | Mark-up                                                    | Description                                           |
|-------------------------------------------------------|------------------------------------------------------------|-------------------------------------------------------|
| [![o_dll]](#- 'Requires external dependencies')       | `[![o_dll]](#- 'Requires external dependencies')`          | Requires external dependencies e.g. DLLs              |
| [![o_inst]](#-)                                       | `[![o_inst]](#-)`                                          | Requires installation                                 |
| [![o_32]](#-)                                         | `[![o_32]](#-)`                                            | 32-bit only/VB6 only                                  |
| [![o_paid]](#- 'Link includes/leads to paid content') | `[![o_paid]](#- 'Link includes/leads to paid content')`    | Link includes/leads to paid content                   |
| [![o_pass]](#-)                                       | `[![o_pass]](#-)`                                          | VBA source code is password protected and/or hidden.  |

Tooltips: append a custom title after the image to give extra detail, e.g.

```md
[![o_dll]](#- 'Requires WinHTTP')
```

#### 4. Github star count

If your repo is a github repo, please also add the star count to your submission. This should follow immediately after your symbology before your title. The syntax to be used is as follows:

```
![GHStars](https://img.shields.io/github/stars/<user-or-org>/<repo>?style&logo=github&label)
```

E.G.

```
- [![p_win]](#-) [![p_mac]](#-) [![a_all]](#-) ![GHStars](https://img.shields.io/github/stars/VBA-tools/VBA-XML?style&logo=github&label) [VBA-XML](https://github.com/VBA-tools/VBA-XML) - XML conversion and parsing.
```

#### Symbology Examples

| Example                                                                     | Markup                                                                       | Description                                                                    |
|-----------------------------------------------------------------------------|------------------------------------------------------------------------------|--------------------------------------------------------------------------------|
| [![p_win]](#-) [![p_mac]](#-) [![a_all]](#-)                                | `[![p_win]](#-) [![p_mac]](#-) [![a_all]](#-)`                               | Compatible on all operating systems and in all applications                    |
| [![p_win]](#-) [![p_nom]](#-) [![a_wd]](#-) [![a_xl]](#-)                   | `[![p_win]](#-) [![p_nom]](#-) [![a_wd]](#-) [![a_xl]](#-)`                  | Only compatible on windows and only works in Word and Excel.                   |
| [![p_win]](#-) [![p_nom]](#-) [![a_xl]](#-) [![o_inst]](#- 'Register OCX')  | `[![p_win]](#-) [![p_nom]](#-) [![a_xl]](#-) [![o_inst]](#- 'Register OCX')` | Only compatible on windows, only works in Excel and requires OCX registration. |
| [![p_win]](#-) [![p_mac]](#-) [![a_all]](#-) [![o_paid]](#- 'One off £200') | `[![p_win]](#-) [![p_mac]](#-) [![a_all]](#-) [![o_paid]](#- 'One off £200')`| Compatible on mac & windows, and in all applications, requires one off £200 license |


### Contribution examples

```
- [![p_win]](#-) [![p_mac]](#-) [![a_all]](#-) ![GHStars](https://img.shields.io/github/stars/sancarn/stdVBA?style&logo=github&label) [stdVBA](https://github.com/sancarn/stdVBA) – Framework of common utilities & collections.
- [![p_win]](#-) [![p_nom]](#-) [![a_xl]](#-)  ![GHStars](https://img.shields.io/github/stars/cristianbuse/VBA-FastJSON?style&logo=github&label) [VBA‑FastJSON](https://github.com/cristianbuse/VBA-FastJSON) – Simple JSON parser for Excel & Access.
- [![p_now]](#-) [![p_mac]](#-) [![a_misc]](#- 'AutoCAD') [AutoCAD‑VBA‑Tools](https://example.com) – Helpers for scripting AutoCAD.
- [![p_win]](#-) [![p_nom]](#-) [![a_misc]](#- 'VB6') [![o_32]](#-) [VB6‑CollectionPlus](https://example.com) – Drop‑in `Collection` with LINQ‑like helpers (VB6‑only).
```

These will render as follows:

- [![p_win]](#-) [![p_mac]](#-) [![a_all]](#-) ![GHStars](https://img.shields.io/github/stars/sancarn/stdVBA?style&logo=github&label) [stdVBA](https://github.com/sancarn/stdVBA) – Framework of common utilities & collections.
- [![p_win]](#-) [![p_nom]](#-) [![a_xl]](#-)  ![GHStars](https://img.shields.io/github/stars/cristianbuse/VBA-FastJSON?style&logo=github&label) [VBA‑FastJSON](https://github.com/cristianbuse/VBA-FastJSON) – Simple JSON parser for Excel & Access.
- [![p_now]](#-) [![p_mac]](#-) [![a_misc]](#- 'AutoCAD') [AutoCAD‑VBA‑Tools](https://example.com) – Helpers for scripting AutoCAD.
- [![p_win]](#-) [![p_nom]](#-) [![a_misc]](#- 'VB6') [![o_32]](#-) [VB6‑CollectionPlus](https://example.com) – Drop‑in `Collection` with LINQ‑like helpers (VB6‑only).

### VB6 Quickstart

While the list is primarily VBA‑oriented, classic VB6 libraries and tools are welcome. Here is a quick guide:

| Icons                               | Markup                                                              | Condition |
|-------------------------------------|---------------------------------------------------------------------|------------------------------------------------------|
| ![p_win] ![p_nom] ![a_all] ![o_32]  | `[![p_win]](#-) [![p_nom]](#-) [![a_all]](#-) [![o_32]](#-)`        | Pure VBA6/VB6 project, with no 64 bit compatibility, no usage of forms (i.e. EB only, no ruby), no OCXs or DLLs. |
| ![p_win] ![p_nom] ![a_all]          | `[![p_win]](#-) [![p_nom]](#-) [![a_all]](#-)`                      | As above but with explicit 64 bit compatibility.     |
| ![p_win] ![p_nom] ![a_misc] ![o_32] | `[![p_win]](#-) [![p_nom]](#-) [![a_misc]](#- 'VB6') [![o_32]](#-)` | Pure VB6 project, with usage of forms (EB+Ruby) but no 64 bit compatibility |

If the project uses external DLLs etc; not accessible by default on a fresh install include the following flag:

```md
[![o_dll]](#- 'Requires external DLL')
```

If the project uses external OCXs which need registering, and/or requires installation of any software, include the following flag:

```md
[![o_inst]](#- 'Requires OCX')
```

If the project comes with paid content for the developer/user, include the following flag:

```md
[![o_paid]](#- '£100pa dev license')
```




<!-- Linker -->

[p_win]: ./resources/WindowsLogo.svg 'Windows'
[p_mac]: ./resources/AppleLogo.svg 'Mac'
[p_now]: ./resources/NotApplicable.svg 'Not Windows'
[p_nom]: ./resources/NotApplicable.svg 'Not Mac'

[a_all]: ./resources/OfficeLogoPlus.svg 'All applications'
[a_wd]: ./resources/WordLogo.svg 'Word'
[a_xl]: ./resources/ExcelLogo.svg 'Excel'
[a_ac]: ./resources/AccessLogo.svg 'Access'
[a_ol]: ./resources/OutlookLogo.svg 'Outlook'
[a_pp]: ./resources/PowerPointLogo.svg 'PowerPoint'
[a_misc]: ./resources/Duck.svg

[o_32]: ./resources/32-Bit.svg '32-bit only'
[o_pass]: ./resources/Padlock.svg 'VBA is password protected'
[o_dll]: ./resources/Dependencies.svg
[o_inst]: ./resources/Installation.svg 'Requires installation'
[o_paid]: ./resources/Money.svg