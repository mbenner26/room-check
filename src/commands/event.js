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

    item.notificationMessages.replaceAsync(
        "roomCheckNotice",
        {
            type:
                Office.MailboxEnums
                    .ItemNotificationMessageType
                    .InformationalMessage,

            message:
                "Bitte Meeting-Anforderungen festlegen.",

            icon: "icon-64",

            persistent: true
        },
        () => {

            event.completed();

        }
    );
}
