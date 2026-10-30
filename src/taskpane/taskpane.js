/* global Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("run").onclick = saveRequirements;
    
    // WICHTIG: Erst speichern, dann Daten laden
    ensureItemIsSaved().then(() => {
      loadAppointmentData();
      loadSavedProperties();
    });
    
    setupConditionalLogic();
  }
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
  if (!item) return;

  try {
    const start = item.start;
    if (start && start instanceof Date && !isNaN(start)) {
      const dateObj = new Date(start);
      document.getElementById("auto-date").value = dateObj.toLocaleDateString('de-DE');
      document.getElementById("auto-time").value = dateObj.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
    }

    document.getElementById("auto-room").value = (typeof item.location === 'string') ? item.location : "Nicht angegeben";
    const host = item.from;
    document.getElementById("auto-host").value = (host && host.displayName) ? host.displayName : "Unbekannt";
  } catch (e) {
    console.error("Fehler beim Laden der Termindaten:", e);
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
        } catch (e) { console.error("Parse Error:", e); }
      }
    }
  });
}

async function saveRequirements(event) {
  event.preventDefault();
  const getCheckedValues = (name) =>
    Array.from(document.querySelectorAll(`input[name="${name}"]:checked`)).map(cb => cb.value);

  const formData = {
    date: document.getElementById("auto-date").value,
    time: document.getElementById("auto-time").value,
    room: document.getElementById("auto-room").value,
    host: document.getElementById("auto-host").value,
    project: document.getElementById("project-name").value,
    representative: document.getElementById("representative").value,
    guestCount: document.getElementById("guest-count").value,
    other: document.getElementById("other-info").value,
    guestStatus: getCheckedValues("guestStatus"),
    catering: getCheckedValues("cateringType"),
    canteen: {
      requested: document.getElementById("canteenRequested").checked,
      date: document.getElementById("cant-date").value,
      time: document.getElementById("cant-time").value,
      guests: document.getElementById("cant-guests").value,
      allergies: document.getElementById("cant-allergies").value
    }
  };

  if (!formData.project || !formData.guestCount) {
    alert("Bitte füllen Sie mindestens das Projekt und die Gästeanzahl aus!");
    return;
  }

  const item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync((loadResult) => {
    if (loadResult.status === Office.AsyncResultStatus.Succeeded) {
      const props = loadResult.value;
      props.set("savedRequirements", JSON.stringify(formData));
      props.saveAsync((saveResult) => {
        if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
          alert("Erfolg: Die Angaben wurden im Termin gespeichert!");
        } else {
          alert("Fehler beim Speichern: " + saveResult.error.message);
        }
      });
    }
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
