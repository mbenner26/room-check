/* global Office */
Office.onReady((info) => {
    if (info.host !== Office.HostType.Outlook) {
        return;
    }

    document
        .getElementById("run")
        .addEventListener("click", saveRequirements);

        loadAppointmentData();
        loadSavedProperties();
        loadAttendees();
        addServiceAttendees();
        setupConditionalLogic();
});


/**
 * Stellt sicher, dass der Termin im Postfach existiert (ID hat),
 * damit CustomProperties gespeichert werden können.
 */
async function ensureItemIsSaved() {
  const item = Office.context.mailbox.item;
  // Wenn itemId null oder undefined ist, muss das Element gespeichert werden
  if (!item.itemId) {
    console.log("Termin hat noch keine ID. Speichere Entwurf...");
    return new Promise((resolve) => {
      item.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Entwurf erfolgreich gespeichert.");
        } else {
          console.error("Fehler beim Speichern des Entwurfs:", result.error);
        }
        resolve(); // Wir machen weiter, auch wenn es fehlschlug, um die UI nicht zu blockieren
      });
    });
  }
  return Promise.resolve();
}

async function loadAppointmentData() {
    const item = Office.context.mailbox.item;
    if (!item) {
        return;
    }
    try {
      if (
          item.start &&
          typeof item.start.getAsync === "function"
        ) {

          item.start.getAsync((result) => {

              if (
                  result.status ===
                  Office.AsyncResultStatus.Succeeded
              ) {

                  const dateObj =
                      new Date(result.value);

                  document.getElementById(
                      "auto-date"
                  ).value =
                      dateObj.toLocaleDateString(
                          "de-DE"
                      );

                  document.getElementById(
                      "auto-time"
                  ).value =
                      dateObj.toLocaleTimeString(
                          "de-DE",
                          {
                              hour: "2-digit",
                              minute: "2-digit"
                          }
                      );
              }
          });
        }
        document.getElementById("auto-host").value =
            Office.context.mailbox.userProfile.displayName || "";
        if (item.location && typeof item.location.getAsync === "function") {
            item.location.getAsync((result) => {
                if (
                    result.status === Office.AsyncResultStatus.Succeeded
                ) {
                    document.getElementById("auto-room").value =
                        result.value || "";
                }
            });
        } else if (typeof item.location === "string") {
            document.getElementById("auto-room").value =
                item.location || "";
        }
    } catch (err) {
        console.error(
            "Fehler beim Laden der Kalenderdaten",
            err
        );
    }
}

async function loadSavedProperties() {
  const item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const props = result.value;
      const savedData = props.get("savedRequirements");
      if (savedData) {
        try {
          const data = JSON.parse(savedData);
          document.getElementById("presentation").value = data.presentation || "";
          document.getElementById("externalGuests").value = data.externalGuests || "";
          document.getElementById("guestDevice").value = data.guestDevice || "";
          document.getElementById("meetingType").value = data.meetingType || "";
          if (data.cateringRequested) {
              document.querySelector(
                  'input[name="cateringDecision"][value="yes"]'
              ).checked = true;
          } else {
              document.querySelector(
                  'input[name="cateringDecision"][value="no"]'
              ).checked = true;
          }
          document.getElementById("project-name").value = data.project || "";
          document.getElementById("representative").value = data.representative || "";
          document.getElementById("guest-count").value = data.guestCount || "";
          document.getElementById("other-info").value = data.other || "";
          document.getElementById("attendee-list").value = data.attendees || "";

          if (data.guestStatus) {
            document.querySelectorAll('input[name="guestStatus"]').forEach(cb => {
              cb.checked = data.guestStatus.includes(cb.value);
            });
          }
          if (data.catering) {
            document.querySelectorAll('input[name="cateringType"]').forEach(cb => {
              cb.checked = data.catering.includes(cb.value);
            });
          }
          if (data.canteen) {
            document.getElementById("canteenRequested").checked = data.canteen.requested;
            document.getElementById("cant-date").value = data.canteen.date || "";
            document.getElementById("cant-time").value = data.canteen.time || "";
            document.getElementById("cant-guests").value = data.canteen.guests || "";
            document.getElementById("cant-allergies").value = data.canteen.allergies || "";
            if (data.canteen.requested) {
              document.getElementById("canteenDetailsContainer").style.display = "grid";
            }
          }
          if (data.externalGuests === "yes") {
                document.getElementById(
                    "guestDeviceContainer"
                ).style.display = "block";
            }
            if (data.cateringRequested) {
                document.getElementById(
                    "cateringDetailsContainer"
                ).style.display = "block";
          }
        } catch (e) { console.error("Parse Error:", e); }
      }
    }
  });
}

function addServiceAttendees() {

    const item = Office.context.mailbox.item;

    if (
        !item ||
        !item.optionalAttendees ||
        !item.optionalAttendees.addAsync
    ) {
        return;
    }

    item.optionalAttendees.getAsync((result) => {

        if (
            result.status !==
            Office.AsyncResultStatus.Succeeded
        ) {
            return;
        }

        const existing =
            result.value.map(
                x => x.emailAddress.toLowerCase()
            );

        const toAdd = [];

        if (
            !existing.includes(
                "email1@domain.de"
            )
        ) {
            toAdd.push({
                displayName:
                    "Raumbuchung",
                emailAddress:
                    "email1@domain.de"
            });
        }

        if (
            !existing.includes(
                "email2@domain.de"
            )
        ) {
            toAdd.push({
                displayName:
                    "Raumbuchung",
                emailAddress:
                    "email2@domain.de"
            });
        }

        if (toAdd.length > 0) {

            item.optionalAttendees.addAsync(
                toAdd
            );

        }

    });
}

function showError(text) {

    Office.context.mailbox.item
        .notificationMessages
        .replaceAsync(
            "validation",
            {
                type:
                    Office.MailboxEnums
                        .ItemNotificationMessageType
                        .ErrorMessage,

                message: text
            }
        );
}

async function saveRequirements() {

    console.log("Speichern gestartet");

    if (!Office.context.mailbox.item) {

        console.error("Kein Outlook Item vorhanden");

        return;
    }

    await ensureItemIsSaved();

    const getCheckedValues = (name) =>
        Array.from(
            document.querySelectorAll(
                `input[name="${name}"]:checked`
            )
        ).map(cb => cb.value);

    const formData = {

        supportRequired:
            document.querySelector(
                'input[name="supportRequired"]:checked'
            )?.value || "no",

        cateringDecision:
            document.querySelector(
                'input[name="cateringDecision"]:checked'
            )?.value || "no",

        presentation:
            document.getElementById("presentation").value,

        externalGuests:
            document.getElementById("externalGuests").value,

        guestDevice:
            document.getElementById("guestDevice").value,

        meetingType:
            document.getElementById("meetingType").value,

        cateringRequested:
            (document.querySelector('input[name="cateringDecision"]:checked')?.value === "yes"),

        date:
            document.getElementById("auto-date").value,

        time:
            document.getElementById("auto-time").value,

        room:
            document.getElementById("auto-room").value,

        host:
            document.getElementById("auto-host").value,

        project:
            document.getElementById("project-name").value,

        representative:
            document.getElementById("representative").value,

        guestCount:
            document.getElementById("guest-count").value,

        other:
            document.getElementById("other-info").value,

        guestStatus:
            getCheckedValues("guestStatus"),

        catering:
            getCheckedValues("cateringType"),

        attendees:
            document.getElementById("attendee-list").value,

        canteen: {
            requested:
                document.getElementById("canteenRequested").checked,

            date:
                document.getElementById("cant-date").value,

            time:
                document.getElementById("cant-time").value,

            guests:
                document.getElementById("cant-guests").value,

            allergies:
                document.getElementById("cant-allergies").value
        }
    };

    const supportDecision =
        document.querySelector(
            'input[name="supportRequired"]:checked'
        );

    const cateringDecision =
        document.querySelector(
            'input[name="cateringDecision"]:checked'
        );

    if (!supportDecision) {

        showError(
            "Bitte auswählen, ob Support benötigt wird."
        );

        return;
    }

    if (!cateringDecision) {

        showError(
            "Bitte auswählen, ob Bewirtung benötigt wird."
        );

        return;
    }

    const missingFields = [];

    if (supportDecision.value === "yes") {

        if (!formData.presentation)
            missingFields.push(
                "Präsentation"
            );

        if (!formData.externalGuests)
            missingFields.push(
                "Externe Gäste"
            );

        if (!formData.meetingType)
            missingFields.push(
                "Videokonferenz"
            );
    }

    if (cateringDecision.value === "yes") {

        if (!formData.date)
            missingFields.push("Datum");

        if (!formData.time)
            missingFields.push("Zeit");

        if (!formData.room)
            missingFields.push("Raum");

        if (!formData.host)
            missingFields.push(
                "Gastgeber"
            );

        if (!formData.project)
            missingFields.push(
                "Projekt"
            );

        if (!formData.guestCount)
            missingFields.push(
                "Anzahl Gäste"
            );
    }

    if (missingFields.length > 0) {

        Office.context.mailbox.item
            .notificationMessages
            .replaceAsync(
                "validation",
                {
                    type:
                        Office.MailboxEnums
                            .ItemNotificationMessageType
                            .ErrorMessage,

                    message:
                        "Fehlende Angaben: "
                        + missingFields.join(", ")
                }
            );

        return;
    }


    console.log(formData);

    const item = Office.context.mailbox.item;

    item.loadCustomPropertiesAsync((loadResult) => {

        if (
            loadResult.status !==
            Office.AsyncResultStatus.Succeeded
        ) {

            console.error(loadResult.error);

            return;
        }

        const props = loadResult.value;

        props.set(
            "savedRequirements",
            JSON.stringify(formData)
        );

        props.saveAsync((saveResult) => {

            if (
                saveResult.status ===
                Office.AsyncResultStatus.Succeeded
            ) {

                console.log(
                    "CustomProperties gespeichert"
                );

                setTimeout(() => {

                    updateAppointmentBody(
                        formData
                    );

                }, 500);

                Office.context.mailbox.item
                    .notificationMessages
                    .replaceAsync(
                        "saved",
                        {
                            type:
                                Office.MailboxEnums
                                    .ItemNotificationMessageType
                                    .InformationalMessage,

                            message:
                                "Raumbuchung erfolgreich gespeichert.",

                            persistent: false
                        }
                    );

            } else {

                console.error(
                    "Speicherfehler:",
                    saveResult.error
                );

            }
        });
    });
}

function setupConditionalLogic() {
    document
        .querySelectorAll(
            'input[name="supportRequired"]'
        )
        .forEach(radio => {
            radio.addEventListener(
                "change",
                () => {
                    document
                        .getElementById(
                            "supportContainer"
                        )
                        .style.display =
                            (
                                radio.value === "yes"
                                &&
                                radio.checked
                            )
                                ? "block"
                                : "none";
                }
            );
        });

    document
        .querySelectorAll(
            'input[name="cateringDecision"]'
        )
        .forEach(radio => {
            radio.addEventListener(
                "change",
                () => {
                    document
                        .getElementById(
                            "cateringDetailsContainer"
                        )
                        .style.display =
                            (
                                radio.value === "yes"
                                &&
                                radio.checked
                            )
                                ? "block"
                                : "none";
                }
            );
        });

    const canteenReq =
        document.getElementById(
            "canteenRequested"
        );
    const canteenDetails =
        document.getElementById(
            "canteenDetailsContainer"
        );
    if (
        canteenReq &&
        canteenDetails
    ) {
        canteenReq.addEventListener(
            "change",
            () => {
                canteenDetails.style.display =
                    canteenReq.checked
                        ? "grid"
                        : "none";
            }
        );
    }

    const externalGuestsSelect =
        document.getElementById(
            "externalGuests"
        );
    const guestDeviceContainer =
        document.getElementById(
            "guestDeviceContainer"
        );
    if (
        externalGuestsSelect &&
        guestDeviceContainer
    ) {
        externalGuestsSelect.addEventListener(
            "change",
            () => {
                guestDeviceContainer.style.display =
                    externalGuestsSelect.value
                        === "yes"
                    ? "block"
                    : "none";
            }
        );
    }
}

function loadAttendees() {

    const item =
        Office.context.mailbox.item;

    item.requiredAttendees.getAsync(
        (result) => {

            if (
                result.status !==
                Office.AsyncResultStatus.Succeeded
            ) {
                return;
            }

            const attendees =
                result.value
                    .map(
                        a => a.displayName
                    );

            document
                .getElementById(
                    "attendee-list"
                )
                .value =
                    attendees.join("\n");
        }
    );
}

function buildBookingText(data) {

    if (
        data.supportRequired === "no"
        &&
        data.cateringDecision === "no"
    ) {

        return `
    <div id="roomcheck-block">

    <table
        width="100%"
        cellpadding="0"
        cellspacing="0"
        style="
            border-collapse:collapse;
            font-family:'Segoe UI', Arial, sans-serif;
            border:1px solid #d1d1d1;
        ">

        <tr>
            <td
                style="
                    background:#0078D4;
                    color:white;
                    padding:15px;
                    font-size:20px;
                    font-weight:bold;
                ">
                Raumbuchung
            </td>
        </tr>

        <tr>
            <td
                style="
                    padding:15px;
                    background:#f8f8f8;
                ">
                Kein Support und keine Bewirtung gewünscht.
            </td>
        </tr>

    </table>

    </div>
    `;
    }


    const guestStatus =
        data.guestStatus?.length > 0
            ? data.guestStatus.join(", ")
            : "-";

    const catering =
        data.catering?.length > 0
            ? data.catering.join(", ")
            : "-";

    return `

<div id="roomcheck-block">

<table
    width="100%"
    cellpadding="0"
    cellspacing="0"
    style="
        border-collapse:collapse;
        font-family:GrotesqueDCA-R;
        border:1px solid #d1d1d1;
        margin-bottom:20px;
    ">

    <tr>
        <td
            colspan="2"
            style="
                background:#0078D4;
                color:white;
                padding:15px;
                font-size:22px;
                font-weight:bold;
            ">
            Raumbuchung
        </td>
    </tr>

    <tr>
        <td
            colspan="2"
            style="
                background:#f3f9fd;
                padding:10px;
                font-weight:bold;
                color:#0078D4;
                border-bottom:1px solid #d1d1d1;
            ">
            Meeting
        </td>
    </tr>

    <tr>
        <td style="padding:8px;font-weight:bold;">Präsentation</td>
        <td style="padding:8px;">${data.presentation || "-"}</td>
    </tr>

    <tr>
        <td style="padding:8px;font-weight:bold;">Videokonferenz</td>
        <td style="padding:8px;">${data.meetingType || "-"}</td>
    </tr>

    <tr>
        <td style="padding:8px;font-weight:bold;">Projekt</td>
        <td style="padding:8px;">${data.project || "-"}</td>
    </tr>

    <tr>
        <td style="padding:8px;font-weight:bold;">Gastgeber</td>
        <td style="padding:8px;">${data.host || "-"}</td>
    </tr>

    <tr>
        <td style="padding:8px;font-weight:bold;">Vertreter</td>
        <td style="padding:8px;">${data.representative || "-"}</td>
    </tr>

    <tr>
        <td
            colspan="2"
            style="
                background:#f3f9fd;
                padding:10px;
                font-weight:bold;
                color:#0078D4;
                border-top:1px solid #d1d1d1;
                border-bottom:1px solid #d1d1d1;
            ">
            Teilnehmer
        </td>
    </tr>

    <tr>
        <td style="padding:8px;font-weight:bold;">Externe Gäste</td>
        <td style="padding:8px;">
            ${data.externalGuests === "yes" ? "Ja" : "Nein"}
        </td>
    </tr>

    <tr>
        <td style="padding:8px;font-weight:bold;">Gerät</td>
        <td style="padding:8px;">
            ${data.guestDevice || "-"}
        </td>
    </tr>

    <tr>
        <td style="padding:8px;font-weight:bold;">Anzahl Gäste</td>
        <td style="padding:8px;">
            ${data.guestCount || "-"}
        </td>
    </tr>

    <tr>
        <td
            colspan="2"
            style="
                background:#f3f9fd;
                padding:10px;
                font-weight:bold;
                color:#0078D4;
                border-top:1px solid #d1d1d1;
                border-bottom:1px solid #d1d1d1;
            ">
            Teilnehmerliste
        </td>
    </tr>

    <tr>
        <td colspan="2" style="padding:10px;">

            ${(data.attendees || "-")
                .replace(/\n/g,"<br>")}

        </td>
    </tr>

    <tr>
        <td style="padding:8px;font-weight:bold;">
            Status
        </td>

        <td style="padding:8px;">
            ${guestStatus}
        </td>
    </tr>

    <tr>
        <td
            colspan="2"
            style="
                background:#f3f9fd;
                padding:10px;
                font-weight:bold;
                color:#0078D4;
                border-top:1px solid #d1d1d1;
                border-bottom:1px solid #d1d1d1;
            ">
            Bewirtung
        </td>
    </tr>

    <tr>
        <td style="padding:8px;font-weight:bold;">Bewirtung</td>
        <td style="padding:8px;">
            ${data.cateringRequested ? "Ja" : "Nein"}
        </td>
    </tr>

    <tr>
        <td style="padding:8px;font-weight:bold;">Art</td>
        <td style="padding:8px;">
            ${catering}
        </td>
    </tr>

    <tr>
        <td
            colspan="2"
            style="
                background:#f3f9fd;
                padding:10px;
                font-weight:bold;
                color:#0078D4;
                border-top:1px solid #d1d1d1;
                border-bottom:1px solid #d1d1d1;
            ">
            Kantinenreservierung
        </td>
    </tr>

    <tr>
        <td style="padding:8px;font-weight:bold;">Reservierung</td>
        <td style="padding:8px;">
            ${data.canteen?.requested ? "Ja" : "Nein"}
        </td>
    </tr>

    <tr>
        <td style="padding:8px;font-weight:bold;">Datum</td>
        <td style="padding:8px;">
            ${data.canteen?.date || "-"}
        </td>
    </tr>

    <tr>
        <td style="padding:8px;font-weight:bold;">Zeit</td>
        <td style="padding:8px;">
            ${data.canteen?.time || "-"}
        </td>
    </tr>

    <tr>
        <td style="padding:8px;font-weight:bold;">Gäste</td>
        <td style="padding:8px;">
            ${data.canteen?.guests || "-"}
        </td>
    </tr>

    <tr>
        <td style="padding:8px;font-weight:bold;">Allergien</td>
        <td style="padding:8px;">
            ${data.canteen?.allergies || "-"}
        </td>
    </tr>

    <tr>
        <td
            colspan="2"
            style="
                background:#f3f9fd;
                padding:10px;
                font-weight:bold;
                color:#0078D4;
                border-top:1px solid #d1d1d1;
                border-bottom:1px solid #d1d1d1;
            ">
            Sonstiges
        </td>
    </tr>

    <tr>
        <td colspan="2" style="padding:10px;">
            ${data.other || "-"}
        </td>
    </tr>

</table>

</div>

`;
}

function updateAppointmentBody(data) {

    const item = Office.context.mailbox.item;

    item.body.getAsync(
        Office.CoercionType.Html,
        (result) => {

            if (
                result.status !==
                Office.AsyncResultStatus.Succeeded
            ) {
                return;
            }

            let body = result.value;

            const parser =
                new DOMParser();

            const doc =
                parser.parseFromString(
                    body,
                    "text/html"
                );

            const existing =
                doc.getElementById(
                    "roomcheck-block"
                );

            if (existing) {
                existing.remove();
            }

            body = doc.body.innerHTML;

            body =
                buildBookingText(data)
                + "<br><br>"
                + body;

                item.body.setAsync(
                    body,
                    {
                        coercionType:
                            Office.CoercionType.Html
                    },
                    (result) => {
                        console.log(
                            "BODY UPDATE STATUS:",
                            result.status
                        );
                        if (
                            result.status !==
                            Office.AsyncResultStatus.Succeeded
                        ) {
                            console.error(
                                "BODY UPDATE ERROR:",
                                result.error
                            );
                        } else {
                            console.log(
                                "BODY UPDATE ERFOLGREICH"
                            );
                        }

                    }
                );
        }
    );
}
