(function () {
  Office.onReady(() => {});

  // Compose butonu
  window.runAttachmentCheck = async (event) => {
    try {
      const item = Office.context.mailbox.item;
      const res = await item.getAttachmentsAsync();
      const list = Array.isArray(res.value) ? res.value : [];
      const hasReal = list.some(a => a && a.isInline === false);

      Office.context.ui.displayDialogAsync(
        "about:blank",
        { height: 30, width: 30, displayInIframe: true },
        () => {}
      );

      console.log("Attachment check:", hasReal ? "VAR" : "YOK");
    } finally {
      event.completed();
    }
  };

  // Gönderirken (OnMessageSend)
  function onMessageSendHandler(event) {
    try {
      Office.context.mailbox.item.getAttachmentsAsync((res) => {
        const hasReal = Array.isArray(res.value) && res.value.some(a => a && a.isInline === false);
        if (hasReal) {
          // UYARI verip gönderimi bir defa engelle
          event.completed({
            allowEvent: false,
            errorMessage: "Bu e-postada ek var. Kontrol edip tekrar gönderin."
          });
        } else {
          event.completed({ allowEvent: true });
        }
      });
    } catch {
      event.completed({ allowEvent: true });
    }
  }
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
})();
