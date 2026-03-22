# Working Manifest — Annotated

This is the production-proven manifest for this Outlook add-in.
Version 1.0.0.5 — confirmed working on New Outlook Mac, New Outlook Windows, Outlook on the web.

## Key decisions captured in comments inline below.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <!-- Generate a fresh GUID for each new add-in: uuidgen in terminal -->
  <Id>D52462A3-A16D-4B1D-AED0-BB3D66DEAC88</Id>
  <Version>1.0.0.5</Version>
  <ProviderName>Your Company</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Auto Signature"/>
  <Description DefaultValue="Automatically injects a remote HTML signature on compose."/>

  <!-- Icons must be publicly accessible PNG files returning HTTP 200.
       Cloudflare Pages: put in /assets/ folder.
       The deployment validator fetches these — a 404 causes silent deployment failure. -->
  <IconUrl DefaultValue="https://your-domain.pages.dev/assets/icon-64.png"/>
  <HighResolutionIconUrl DefaultValue="https://your-domain.pages.dev/assets/icon-128.png"/>
  <SupportUrl DefaultValue="https://your-domain.pages.dev"/>

  <!-- Minimum Outlook requirement for event-based activation -->
  <Requirements>
    <Sets DefaultMinVersion="1.3">
      <Set Name="Mailbox" MinVersion="1.3"/>
    </Sets>
  </Requirements>

  <!-- Required boilerplate for mail apps — not used in event-only mode.
       SourceLocation must return 200 — use extensionless URL (Cloudflare Pretty URLs). -->
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="https://your-domain.pages.dev/commands"/>
        <RequestedHeight>250</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <!-- VersionOverrides is where event-based activation lives.
       Both V1_0 and V1_1 wrappers are required — nested structure is intentional. -->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">

      <Requirements>
        <bt:Sets DefaultMinVersion="1.3">
          <bt:Set Name="Mailbox" MinVersion="1.3"/>
        </bt:Sets>
      </Requirements>

      <Hosts>
        <Host xsi:type="MailHost">

          <!-- RUNTIMES block is REQUIRED for event-based activation.
               Without it you get "resid tag is invalid" at sideload time.
               
               webViewRuntime: HTML page — used by Outlook on the web
               jsRuntime: JS file — used by Outlook desktop (via Override)
               
               resid IDs must be plain camelCase — NO DOTS.
               Dots in resid IDs cause "resid tag is invalid" on some Outlook versions. -->
          <Runtimes>
            <Runtime resid="webViewRuntime">
              <Override type="javascript" resid="jsRuntime"/>
            </Runtime>
          </Runtimes>

          <DesktopFormFactor>

            <!-- DO NOT include <FunctionFile> here for LaunchEvent-only add-ins.
                 FunctionFile is only for task pane / ribbon command add-ins.
                 Including it causes "resid tag is invalid". -->

            <ExtensionPoint xsi:type="LaunchEvent">
              <LaunchEvents>
                <!-- FunctionName must exactly match Office.actions.associate() call in commands.js -->
                <LaunchEvent Type="OnNewMessageCompose" FunctionName="onNewMessageCompose"/>
                <!-- Uncomment to also handle calendar invites:
                <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onNewMessageCompose"/>
                -->
              </LaunchEvents>
              <!-- SourceLocation here references the webViewRuntime resid -->
              <SourceLocation resid="webViewRuntime"/>
            </ExtensionPoint>

          </DesktopFormFactor>
        </Host>
      </Hosts>

      <Resources>
        <bt:Urls>
          <!-- HTML page that loads commands.js — used by Outlook on the web.
               Cloudflare Pages serves commands.html at /commands (extensionless).
               Must return HTTP 200 — no redirects. -->
          <bt:Url id="webViewRuntime" DefaultValue="https://your-domain.pages.dev/commands"/>
          <!-- JS file executed directly by Outlook desktop.
               .js extension is NOT affected by Cloudflare Pretty URLs — returns 200 as-is. -->
          <bt:Url id="jsRuntime" DefaultValue="https://your-domain.pages.dev/commands.js"/>
        </bt:Urls>
        <bt:ShortStrings>
          <bt:String id="addInName" DefaultValue="Auto Signature"/>
        </bt:ShortStrings>
        <bt:LongStrings>
          <bt:String id="addInDescription" DefaultValue="Automatically injects a remote HTML signature on compose."/>
        </bt:LongStrings>
      </Resources>

    </VersionOverrides>
  </VersionOverrides>

</OfficeApp>
```
