/* global Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("run").onclick = saveRequirements;
    
    // 1. Zuerst die harten Fakten aus dem Outlook-Termin laden
    loadAppointmentData();

    // 2. Dann versuchen, eventuelle alte (gespeicherte) Informationen zu laden
    loadSavedProperties();

    setupConditionalLogic();
  }
});

/**
* Lädt die harten Fakt-Daten aus dem Outlook-Element
* (Datum, Zeit, Raum, Gastgeber)
*/
async function loadAppointmentData() {
  const item = Office.context.mailbox.item;
  if (!item) return;

  const start = item.start;
  if (start) {
    const dateObj = new Date(start);
    document.getElementById("auto-date").innerText = dateObj.toLocaleDateString('de-DE');
    document.getElementById("auto-time").innerText = dateObj.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  }

  document.getElementById("auto-room").innerText = item.location || "Nicht angegeben";
  const host = item.from;
  document.getElementById("auto-host").innerText = (host && host.displayName) ? host.displayName : "Unbekannt";
}

/**
* Lädt benutzerdefinierte Daten aus den CustomProperties
*/
async function loadSavedProperties() {
  const item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const props = result.value;
      const savedData = props.get("savedRequirements");

      if (savedData) {
        try {
          const data = JSON.parse(savedData);
          console.log("Geladene Daten aus Speicher:", data);

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
            document.getElementById("cant-time").value = data.canteen.time || ""; // FIX: Tippfehler korrigiert
            document.getElementById("cant-guests").value = data.canteen.guests || "";
            document.getElementById("cant-allergies").value = data.canteen.allergies || "";

            if (data.canteen.requested) {
              document.getElementById("canteenDetailsContainer").style.display = "grid";
            }
          }
        } catch (e) {
          console.error("Fehler beim Parsen der gespeicherten Daten:", e);
        }
      }
    }
  });
}

async function saveRequirements(event) {
  event.preventDefault();

  // FIX: Klammer bei querySelectorAll ergänzt
  const getCheckedValues = (name) =>
    Array.from(document.querySelectorAll(`input[name="${name}"]:checked`)).map(cb => cb.value);

  const formData = {
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
          alert("Fehler beim Speichern der Daten.");
        }
      });
    }
  });
}

function setupConditionalLogic() {
  // Kantine Logik
  const canteenReq = document.getElementById("canteenRequested");
  const canteenDetails = document.getElementById("canteenDetailsContainer");
  if (canteenReq && canteenDetails) {
    canteenReq.addEventListener("change", () => {
      canteenDetails.style.display = canteenReq.checked ? "grid" : "none";
    });
  }
  
  // Bewirtung Haupt-Checkbox Logik (falls vorhanden)
  const cateringRequested = document.getElementById("cateringRequested");
  const cateringDetailsContainer = document.getElementById("cateringDetailsContainer");
  if (cateringRequested && cateringDetailsContainer) {
    cateringRequested.addEventListener("change", () => {
      cateringDetailsContainer.style.display = cateringRequested.checked ? "block" : "none";
    });
  }
}
