<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xsi:type="MailApp">
  <Id>9F2E3A64-7E9D-4E34-9E9C-9B7A1A7E00A1</Id>
  <Version>1.0.2.2</Version>
  <ProviderName>Trendyol</ProviderName>
  <DefaultLocale>tr-TR</DefaultLocale>
  <DisplayName DefaultValue="Test (Pane)"/>
  <Description DefaultValue="Yükleme testi için en basit görev bölmesi."/>

  <!-- İKONLAR -->
  <IconUrl DefaultValue="https://duyguyil09-cell.github.io/outlook-attachment-alert/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://duyguyil09-cell.github.io/outlook-attachment-alert/icon-80.png"/>

  <SupportUrl DefaultValue="https://duyguyil09-cell.github.io/outlook-attachment-alert/home.html"/>

  <AppDomains>
    <AppDomain>duyguyil09-cell.github.io</AppDomain>
  </AppDomains>

  <Hosts>
    <Host Name="Mailbox"/>
  </Hosts>

  <!-- GEREKSİNİM -->
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.1"/>
    </Sets>
  </Requirements>

  <!-- FORM AYARLARI (pane) -->
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="https://duyguyil09-cell.github.io/outlook-attachment-alert/home.html"/>
        <RequestedHeight>300</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>

  <!-- İZİN -->
  <Permissions>ReadItem</Permissions>

  <!-- KURAL -->
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
  </Rule>
</OfficeApp>
