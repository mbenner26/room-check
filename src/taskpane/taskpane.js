/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("run").onclick = saveRequirements;
    
    // Automatische Daten laden
    loadAppointmentData();
    setupConditionalLogic();
  }
});

/**
 * Steuert die Sichtbarkeit der Felder basierend auf den Antworten
 */
function setupConditionalLogic() {
  const externalGuestsSelect = document.getElementById("externalGuests");
  const guestDeviceContainer = document.getElementById("guestDeviceContainer");

  externalGuestsSelect.addEventListener("change", () => {
    if (externalGuestsSelect.value === "yes") {
      guestDeviceContainer.style.display = "block";
    } else {
      guestDeviceContainer.style.display = "none";
      // Wert zurücksetzen, wenn ausgeblendet
      document.getElementById("guestDevice").value = "";
    }
  });

  // NEU: Logik für Bewirtung
  const cateringRequested = document.getElementById("cateringRequested");
  const cateringDetailsContainer = document.getElementById("cateringDetailsContainer");
  
  cateringRequested.addEventListener("change", () => {
    if (cateringRequested.checked) {
      cateringDetailsContainer.style.display = "block";
    } else {
      cateringDetailsContainer.style.display = "none";
      // Optional: Alle Unter-Checkboxen zurücksetzen, wenn Haupt-Checkbox deaktiviert wird
      document.querySelectorAll('input[name="cateringType"]').forEach(cb => cb.checked = false);
    }
  });
}

async function loadAppointmentData() {
  const item = Office.context.mailbox.item;
  
  // In Outlook-Addins sind viele Infos direkt im 'item' Objekt verfügbar
  // Hinweis: Je nach Outlook-Version/Plattform variiert der Zugriff
  
  // Datum und Zeit (Beispielhaft für den Start-Termin)
  const start = item.appointmentStart; 
  if (start) {
    const dateObj = new Date(start);
    document.getElementById("auto-date").innerText = dateObj.toLocaleDateString('de-DE');
    document.getElementById("auto-time").innerText = dateObj.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  }

  // Ort/Raum
  document.getElementById("auto-room").innerText = item.location || "Nicht angegeben";
  
  // Gastgeber (Organisator)
  // In der Regel über den Organizer-Namen des Items
  document.getElementById("auto-host").innerText = item.organizer || "Unbekannt";
}

function setupConditionalLogic() {
  const canteenReq = document.getElementById("canteenRequested");
  const canteenDetails = document.getElementById("canteenDetailsContainer");
  
  canteenReq.addEventListener("change", () => {
    canteenDetails.style.display = canteenReq.checked ? "grid" : "none";
  });
}


/**
 * Sammelt die Daten aus dem Formular und bereitet sie vor
 */
async function saveRequirements(event) {
  event.preventDefault();

  // Sammle Checkboxen
  const getCheckedValues = (name) => 
    Array.from(document.querySelectorAll(`input[name="${name}"]:checked`)).map(cb => cb.value);

  
  // Sammle alle aktivierten Bewirtungs-Optionen in einem Array
  const selectedCatering = Array.from(document.querySelectorAll('input[name="cateringType"]:checked'))
                                .map(cb => cb.value);

  const formData = {
    // Automatische Daten
    date: document.getElementById("auto-date").innerText,
    time: document.getElementById("auto-time").innerText,
    room: document.getElementById("auto-room").innerText,
    host: document.getElementById("auto-host").innerText,
    // Manuelle Daten
    project: document.getElementById("project-name").value,
    representative: document.getElementById("representative").value,
    guestCount: document.getElementById("guest-count").value,
    other: document.getElementById("other-info").value,
    // Status & Bewirtung
    guestStatus: getCheckedValues("guestStatus"),
    catering: getCheckedValues("cateringType"),
    // Kantine
    canteen: {
      requested: document.getElementById("canteenRequested").checked,
      date: document.getElementById("cant-date").value,
      time: document.getElementById("cant-time").value,
      guests: document.getElementById("cant-guests").value,
      allergies: document.getElementById("cant-allergies").value
    }
  };

  // Validierung
  if (!formData.presentation || !formData.externalGuests || !formData.meetingType) {
    alert("Bitte alle Pflichtfelder ausfüllen!");
    return;
  }

  // Validierung Bewirtung: Wenn "Ja", muss mindestens eine Option gewählt sein
  if (formData.cateringRequested && selectedCatering.length === 0) {
    alert("Bitte wählen Sie mindestens eine Art der Bewirtung aus.");
    return;
  }

  console.log("Gesammelte Daten:", formData);
  alert("Daten wurden lokal erfasst!");
}

  // TODO: API Call an Azure Function / Web-Service
  // await postDataToBackend(formData);
