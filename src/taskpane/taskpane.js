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
 * Lädt die harten Fakret-Daten aus dem Outlook-Element
 * (Datum, Zeit, Raum, Gastgeber)
 */
async function loadAppointmentData() {
  const item = Office.context.mailbox.item;
  if (!item) return;

  // Startzeit abgreifen
  const start = item.start;
  if (start) {
    const dateObj = new Date(start);
    document.getElementById("auto-date").innerText = dateObj.toLocaleDateString('de-DE');
    document.getElementById("auto-time").innerText = dateObj.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  }
  
  // Ort und Gastgeber
  document.getElementById("auto-room").innerText = item.location || "Nicht angegeben";
  const host = item.from;
  document.getElementById("auto-host").innerText = (host && host.displayName) ? host.displayName : "Unbekannt";
}

/**
 * Lädt nur die benutzerdefinierten Daten (Projekt, Gäste, etc.)
 * Diese werden NUR geladen, wenn sie jemals per 'Speichern' Button gespeichert wurden.
 */
async function loadSavedProperties() {
  const item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const props = result.value;
      const savedData = props.get("savedRequirements");
      
      if (savedData) {
        const data = JSON.parse(savedData);
        console.log("Geladene Daten aus Speicher:", data);

        // Felder befüllen
        document.getElementById("project-name").value = data.project || "";
        document.getElementById("representative").value = data.representative || "";
        document.getElementById("guest-count").value = data.guestCount || "";
        document.getElementById("other-info").value = data.other || "";
        
        // Status Checkboxen
        if (data.guestStatus) {
          document.querySelectorAll('input[name="guestStatus"]').forEach(cb => {
            cb.checked = data.guestStatus.includes(cb.value);
          });
        }

        // Catering Checkboxen
        if (data.catering) {
          document.querySelectorAll('input[name="cateringType"]').forEach(cb => {
            cb.checked = data.catering.includes(cb.value);
          });
        }

        // Kantine
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
      }
    }
  });
}

async function saveRequirements(event) {
  event.preventDefault();

  const getCheckedValues = (name) =>
    Array.from(document.querySelectorAll(`input[name="${name}"]:checked`)).map(cb => cb.value);

  // Alle Daten sammeln
  const formData = {
    // WICHTIG: Wir speichern hier NICHT die auto-Felder (Datum/Raum), 
    // da diese immer direkt aus dem aktuellen Outlook-Termin kommen sollten!
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

  // Validierung
  if (!formData.project || !formData.guestCount) {
    alert("Bitte füllen Sie das Projekt und die Gästeanzahl aus!");
    return;
  }

  // Speichern in die CustomProperties
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
  // ... (bleibt identisch wie oben)
  const canteenReq = document.getElementById("canteenRequested");
  const canteenDetails = document.getElementById("canteenDetailsContainer");
  if (canteenReq && canteenDetails) {
    canteenReq.addEventListener("change", () => {
      canteenDetails.style.display = canteenReq.checked ? "grid" : "none";
    });
  }
  // ...
}
