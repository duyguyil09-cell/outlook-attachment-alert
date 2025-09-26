(function(){
  function onMessageSendHandler(event){
    try{
      const item = Office.context.mailbox.item;
      item.getAttachmentsAsync((res)=>{
        if (res.status !== Office.AsyncResultStatus.Succeeded){
          return event.completed({ allowEvent: true });
        }
        const list = Array.isArray(res.value) ? res.value : [];
        const hasReal = list.some(a => a && a.isInline === false);
        if (hasReal){
          return event.completed({
            allowEvent: false,
            errorMessage: "Dikkat: Bu e-postada ek(ler) var. Göndermek istiyor musunuz?",
            errorMessageMarkdown: "**Dikkat:** Bu e-postada **ek(ler)** var.\n\nGöndermek istiyor musunuz?"
          });
        }
        event.completed({ allowEvent: true });
      });
    } catch {
      event.completed({ allowEvent: true });
    }
  }
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
})();
