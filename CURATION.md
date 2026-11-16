# Curation guidelines

Awesome VBA celebrates libraries, tools and resources that make working with VBA and classic VB6 more useful, enjoyable or surprising. We want each entry to give readers a clear reason to explore it.

Our audience includes people writing their first macros, developers building and maintaining applications, and those exploring the language's internals. VBA extends beyond Excel: resources for Word, Access, Outlook, PowerPoint and other VBA hosts are welcome, as are classic VB6 libraries and tools.

When suggesting a resource, start with this question: **What makes it exceptional, and why would a VBA or VB6 developer choose it?**

## What belongs here

- **Libraries and frameworks** that provide useful, reusable capabilities or meaningfully improve how developers write and maintain code.
- **Developer tools and add-ins** that improve editing, debugging, testing, version control, deployment or other parts of VBA and VB6 development.
- **Exceptional examples and showcases**, including games, simulations, artwork and finished applications whose VBA or VB6 implementation stands out through exceptional programming, reusable engineering or an instructive approach.
- **Learning resources** that stand out for their clarity, depth, originality or practical value, from introductory lessons to advanced techniques.
- **External tools and integrations** that offer substantial value for working with VBA or VB6, including interoperability with other languages and environments. They do not have to be written in VBA or VB6 themselves.
- **Reference material, style guides and historical resources** that help readers understand the languages, their runtimes or their development practices particularly well.

## Awesome VBA and Awesome Excel

[Awesome Excel](https://github.com/sancarn/awesome-excel) focuses on resources for working in Excel. The inclusion case centres on what a tool, workbook or other resource enables Excel users to do, understand or enjoy. Its implementation may use VBA, formulas or another technology.

Awesome VBA focuses on VBA and VB6 development. Generic, reusable code that can serve projects across different environments is central to this list. For code submissions, explain what makes the implementation exceptional in its own right: its capabilities, design, engineering or techniques that other developers can use or learn from. Developer tools and learning resources qualify through the value they offer people working with these languages.

Portability is valuable, but code does not have to run in multiple hosts to qualify. A resource specific to Excel, Access or another environment can belong here when its implementation offers exceptional value to VBA or VB6 developers. The host application alone does not determine which list it suits.

For example, an outstanding Excel planning workbook may fit Awesome Excel because of the experience it offers spreadsheet users. Its use of ordinary VBA to connect buttons and update cells would not by itself establish a case here. A reusable scheduling library or an exceptional VBA technique within that workbook could provide a separate inclusion case. A resource can qualify for both lists when it offers distinct value to each audience.

Each list has its own quality threshold. Falling below Awesome Excel's threshold does not establish a case for inclusion here. Using VBA or making its source accessible is insufficient on its own; explain the specific value for this audience. General spreadsheet templates, formula collections or unrelated programming tools need a clear VBA or VB6 inclusion case.

Equally, exclusion from Awesome VBA does not rule out inclusion in Awesome Excel or another list. An outstanding spreadsheet may merit inclusion in Awesome Excel even when its VBA implementation offers little of particular interest to VBA developers. Assess each resource against the audience, scope and quality criteria of the list where it is being proposed.

## What makes something awesome

### A clear purpose and audience

Assess a resource against what it promises and who it serves. A library or utility should demonstrate a meaningful practical benefit. A teaching resource should help its audience understand or apply something particularly well. Art and showcases can qualify through creativity or craftsmanship without being practically useful.

A beginner resource can be exceptional through clear explanations and well-chosen examples. An advanced resource can serve a narrow audience exceptionally well. Complexity and broad appeal are not requirements.

A utility with weak practical value needs an independent case for inclusion as a showcase. Recognising the work behind it does not establish that case.

### A reason to choose it

Explain what stands out beyond a feature list. Where built-in language features, host APIs or other resources already address the task, identify the advantage: substantially better usability, automation, reliability, performance, explanation or capabilities. An excellent implementation of a familiar workflow can qualify.

For example, another dictionary implementation or JSON parser may deserve a place because of its API, correctness, portability or performance. Merely implementing a familiar feature does not establish that case. A resource can complement existing entries without replacing them or outperforming them at their own specialties.

Practical constraints matter to this audience. Code that can be imported into a project without installation or extra dependencies can be especially useful in restricted environments. That advantage should support a meaningful capability or a particularly good implementation.

### Quality of the result

Judge the resulting experience, insight or capability. Large amounts of code, months of work, low-level techniques or a manual recreation of a familiar game demonstrate dedication; they do not independently establish exceptional quality.

For libraries, consider whether the API and implementation deliver their promised value and can be used reliably. Clear setup instructions, examples and an understandable dependency model help readers put them to work.

For showcases, explain what stands out beyond the surprise that something runs in VBA or VB6. This could be inventive use of the language or host, thoughtful interaction implemented particularly well, reusable techniques or substantial educational value. A showcase need not be a library or tutorial, but the recommendation here should identify what is exceptional about its VBA or VB6 implementation.

Straightforward techniques can produce exceptional work. Technical mystery, complexity and avoiding external dependencies are not requirements; assess the quality and inventiveness of their use.

### Understanding that carries over

Learning resources should demonstrate how they teach particularly well. Clear explanations, worked examples, meaningful practice and interactive exploration can all contribute.

A consistently excellent VBA or VB6 blog can earn a recommendation through its body of work. An individual post or forum answer should be exceptional in its own right or have an exceptional reason to stand alone in the list. Choose the level of recommendation that best represents the value: the site as a whole or the particular resource.

Code examples should offer insight readers can apply: why an approach works, what trade-offs it makes or how it solves a difficult problem. A code listing or visual breakdown may expose structure without explaining behaviour. More detail or interaction is valuable when it makes ideas easier to understand or apply.

### Enough evidence to recommend it

Provide accessible material that supports the inclusion case: a working library or tool, substantive lessons, source or a convincing demonstration. Free downloads and accessible code make assessment easier; they do not by themselves justify inclusion.

Review claims in proportion to their importance. Correctness is central to a parser; reproducible comparisons matter to a library recommended for speed. Unresolved problems with a resource's main promise can justify deferring or declining it.

Be clear about what was checked. A creator's screenshots, documentation or compatibility claims are different evidence from running code in a particular host. Static inspection can support a review, but should not be presented as gameplay, runtime or compatibility testing. Reviewers need not test every environment; state the limits of the evidence.

### Compatibility and practical requirements

A resource does not need to support every operating system, host or architecture to qualify. A Windows-only library, a host-specific add-in or a historical VB6 example can still be an excellent recommendation for its intended audience.

Describe relevant requirements accurately: Windows or macOS, the host application, VBA or VB6 version, 32-bit or 64-bit support, external libraries, installation and registration steps. Do not infer support for every VBA host from an Excel example, or 64-bit VBA compatibility from working VB6 code.

Use the [README's symbols](README.md#symbology) for supported environments and constraints, with tooltips or a short description for details the icons cannot express, such as a 64-bit-only requirement. Distinguish documented support from personally verified compatibility, and flag important unknowns in the submission.

### Room for improvement

Resources need not be perfect or feature-complete. Assess whether limitations undermine the reason for recommending them. Feature requests, rough edges and ordinary licensing traffic are not automatic grounds for rejection. Consider security and privacy findings on their evidence and impact.

Age or a quiet commit history alone does not disqualify a resource. Stable libraries and historical material can remain valuable; describe relevant limitations and prefer a reliable source or archive when an original link disappears.

Keep the merit decision separate from improvement feedback. Fixing access, compatibility or implementation problems may allow reconsideration, but does not automatically make a resource awesome.

## Suggesting a resource

Open an issue or submit a pull request with:

- A direct link to the resource or its original creator. Prefer the relevant repository, release, documentation or forum post over a general promotional page, and omit referral parameters.
- A short, factual description of what it does and a specific explanation of what makes it worth including.
- Its intended audience and, where useful, why someone would choose it over relevant alternatives.
- Enough material to assess it, with any limits on what you have personally verified.
- Significant costs, access restrictions, compatibility requirements and dependencies, such as external DLLs or OCXs, installation, accounts, external services or a particular host version.

Commercial resources are welcome. Identify paid content and password-protected or hidden source using the existing symbols where applicable. Source access is especially helpful when recommending an implementation for study or reuse; a tool can earn a place through its value to developers without publishing its source.

Follow the [contribution guidelines](Contributing.md) for entry formatting and symbology. Keep listing descriptions concise, accurate and supported by the resource; avoid universal claims such as "works in all VBA hosts" unless they can be substantiated. Put the fuller inclusion case and review evidence in the issue or pull request.

Adapted from [Awesome Excel's curation guidelines](https://github.com/sancarn/awesome-excel/blob/main/CURATION.md).
