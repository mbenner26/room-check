Office.onReady(() => {
  Office.actions.associate(
    "onNewAppointmentOrganizer",
    onNewAppointmentOrganizer
  );
});

function onNewAppointmentOrganizer(event) {
  Office.context.mailbox.item.notificationMessages.addAsync(
    "roomCheckNotice",
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Bitte Meeting-Anforderungen festlegen.",
      icon: "icon-64",
      persistent: true
    },
    () => {
      event.completed();
    }
  );
}
