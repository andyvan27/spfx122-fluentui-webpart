# spfx-122-fluentui-webpart

## Summary

SampleSPFX 1.22 webpart using FluentUI, React TypeScript, PnPJs, PnPGraph.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.22.0-green.svg)

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - `npm install -g @rushstack/heft`
  - `npm install`
  - `heft start`

## From scratch
- Install and use node 22:
```
nvm install 22
nvm alias default 22
```

- Install SharePoint Framework ToolKit VS Code extension
- Install Yeoman and related
```
npm install -g yo
npm install -g @rushstack/heft
npm install -g @microsoft/generator-sharepoint@1.22.0
```
- Use SPFX VS Code extension to create new SPFX WebPart with all the options
- Also on the extension UI after created, run those common tasks: 
  - Trust self-signed dev cert (`heft trust-dev-cert`)
  - Test (`heft test`)
  - Start (`heft start`)

## Install local workbench
```
npm install @microsoft/sp-webpart-workbench@1.22.0 --save-dev
```

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [Heft Documentation](https://heft.rushstack.io/)