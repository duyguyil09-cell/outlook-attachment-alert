(function () {
  function onMessageSendHandler(event) {
    try {
      const item = Office.context.mailbox.item;
      item.getAttachmentsAsync((res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          return event.completed({ allowEvent: true });
        }
        const list = Array.isArray(res.value) ? res.value : [];
        const hasRealAttachments = list.some(a => a && a.isInline === false);

        if (hasRealAttachments) {
          // Gönderimi bir kez engelle; kullanıcı mesajı görsün.
          return event.completed({
            allowEvent: false,
            errorMessage:
              "Dikkat: Bu e-postada ek(ler) var. Göndermeden önce tekrar kontrol etmek ister misiniz? (Devam etmek için uyarıyı kapatıp tekrar Gönder'e basın.)",
            errorMessageMarkdown:
              "**Dikkat:** Bu e-postada **ek(ler)** var.\n\nDevam etmek için uyarıyı kapatıp **Gönder**'e tekrar basın."
          });
        }
        // Ek yoksa gönder
        event.completed({ allowEvent: true });
      });
    } catch {
      // Herhangi bir hata olursa gönderimi engelleme
      event.completed({ allowEvent: true });
    }
  }

  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
})();
