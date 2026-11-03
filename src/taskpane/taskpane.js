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
        if (item.start instanceof Date) {
            const dateObj = new Date(item.start);
            document.getElementById("auto-date").value =
                dateObj.toLocaleDateString("de-DE");
            document.getElementById("auto-time").value =
                dateObj.toLocaleTimeString("de-DE", {
                    hour: "2-digit",
                    minute: "2-digit"
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
          document.getElementById("cateringRequested").checked = data.cateringRequested || false;
          document.getElementById("project-name").value = data.project || "";
          document.getElementById("representative").value = data.representative || "";
          document.getElementById("guest-count").value = data.guestCount || "";
          document.getElementById("other-info").value = data.other || "";

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

        presentation:
            document.getElementById("presentation").value,

        externalGuests:
            document.getElementById("externalGuests").value,

        guestDevice:
            document.getElementById("guestDevice").value,

        meetingType:
            document.getElementById("meetingType").value,

        cateringRequested:
            document.getElementById("cateringRequested").checked,

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

    const missingFields = [];

    if (!formData.presentation)
        missingFields.push("Präsentation");

    if (!formData.externalGuests)
        missingFields.push("Externe Gäste");

    if (!formData.meetingType)
        missingFields.push("Videokonferenz");

    if (!formData.project)
        missingFields.push("Projekt");

    if (!formData.guestCount)
        missingFields.push("Anzahl Gäste");

    if (missingFields.length > 0) {

        Office.context.mailbox.item.notificationMessages.replaceAsync(
            "validation",
            {
                type:
                    Office.MailboxEnums
                        .ItemNotificationMessageType
                        .ErrorMessage,

                message:
                    "Pflichtfelder fehlen: "
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

                updateAppointmentBody(formData);

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
  const externalGuestsSelect = document.getElementById("externalGuests");
  const guestDeviceContainer = document.getElementById("guestDeviceContainer");
  if (externalGuestsSelect && guestDeviceContainer) {
    externalGuestsSelect.addEventListener("change", () => {
      guestDeviceContainer.style.display = (externalGuestsSelect.value === "yes" || externalGuestsSelect.value === "Ja") ? "block" : "none";
    });
  }
  const canteenReq = document.getElementById("canteenRequested");
  const canteenDetails = document.getElementById("canteenDetailsContainer");
  if (canteenReq && canteenDetails) {
    canteenReq.addEventListener("change", () => {
      canteenDetails.style.display = canteenReq.checked ? "grid" : "none";
    });
  }
  const cateringRequested = document.getElementById("cateringRequested");
  const cateringDetailsContainer = document.getElementById("cateringDetailsContainer");
  if (cateringRequested && cateringDetailsContainer) {
    cateringRequested.addEventListener("change", () => {
      cateringDetailsContainer.style.display = cateringRequested.checked ? "block" : "none";
    });
  }
}

function buildBookingText(data) {

    const guestStatus =
        (data.guestStatus || []).join(", ");

    const catering =
        (data.catering || []).join(", ");

    return `
<hr>

<h2>Raumbuchung</h2>

<p>
<b>Präsentation:</b> ${data.presentation || "-"}<br>
<b>Externe Gäste:</b> ${data.externalGuests || "-"}<br>
<b>Gerät externe Gäste:</b> ${data.guestDevice || "-"}<br>
<b>Videokonferenz:</b> ${data.meetingType || "-"}
</p>

<p>
<b>Projekt:</b> ${data.project || "-"}<br>
<b>Gastgeber:</b> ${data.host || "-"}<br>
<b>Vertreter:</b> ${data.representative || "-"}<br>
<b>Anzahl Gäste:</b> ${data.guestCount || "-"}
</p>

<p>
<b>Gästestatus:</b> ${guestStatus || "-"}
</p>

<p>
<b>Bewirtung:</b> ${catering || "-"}
</p>

<p>
<b>Kantinenreservierung:</b>
${data.canteen?.requested ? "Ja" : "Nein"}
</p>

<p>
<b>Datum:</b> ${data.canteen?.date || "-"}<br>
<b>Zeit:</b> ${data.canteen?.time || "-"}<br>
<b>Gäste:</b> ${data.canteen?.guests || "-"}<br>
<b>Allergien:</b> ${data.canteen?.allergies || "-"}
</p>

<p>
<b>Sonstiges:</b><br>
${data.other || "-"}
</p>

<!-- ROOMCHECK_START -->
${encodeURIComponent(JSON.stringify(data))}
<!-- ROOMCHECK_END -->
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

            body = body.replace(
                /<!-- ROOMCHECK_START -->[\s\S]*?<!-- ROOMCHECK_END -->/g,
                ""
            );

            body =
                buildBookingText(data)
                + "<br><br>"
                + body;

            item.body.setAsync(
                body,
                {
                    coercionType:
                        Office.CoercionType.Html
                }
            );
        }
    );
}
