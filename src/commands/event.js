/* global Office */

Office.onReady(() => {
    Office.actions.associate(
        "onNewAppointmentOrganizer",
        onNewAppointmentOrganizer
    );
});

function onNewAppointmentOrganizer(event) {
    const item = Office.context.mailbox.item;

    if (!item) {
        event.completed();
        return;
    }

    item.notificationMessages.addAsync(
        "roomCheckNotice",
        {
            type: Office.MailboxEnums.ItemNotificationMessageType.WarningMessage,
            message: "Bitte Meeting-Anforderungen festlegen.",
            icon: "icon-64",
            persistent: true
        },
        (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                console.error(
                    "Die Benachrichtigung konnte nicht angezeigt werden:",
                    result.error
                );
            }

            event.completed();
        }
    );
}
