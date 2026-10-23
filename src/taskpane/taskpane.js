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
* Steuert die Sichtbarkeit ALLER bedingten Felder
*/
function setupConditionalLogic() {
  // 1. Externe Gäste / Gerät (falls noch in HTML vorhanden)
  const externalGuestsSelect = document.getElementById("externalGuests");
  const guestDeviceContainer = document.getElementById("guestDeviceContainer");
  if (externalGuestsSelect && guestDeviceContainer) {
    externalGuestsSelect.addEventListener("change", () => {
      guestDeviceContainer.style.display = (externalGuestsSelect.value === "yes") ? "block" : "none";
    });
  }

  // 2. Kantine / Details
  const canteenReq = document.getElementById("canteenRequested");
  const canteenDetails = document.getElementById("canteenDetailsContainer");
  if (canteenReq && canteenDetails) {
    canteenReq.addEventListener("change", () => {
      canteenDetails.style.display = canteenReq.checked ? "grid" : "none";
    });
  }
  
  // 3. Bewirtung (falls du hier eine Haupt-Checkbox hast)
  const cateringRequested = document.getElementById("cateringRequested");
  const cateringDetailsContainer = document.getElementById("cateringDetailsContainer");
  if (cateringRequested && cateringDetailsContainer) {
    cateringRequested.addEventListener("change", () => {
      cateringDetailsContainer.style.display = cateringRequested.checked ? "block" : "none";
    });
  }
}

async function loadAppointmentData() {
  const item = Office.context.mailbox.item;
  if (!item) return;

  const start = item.appointmentStart;
  if (start) {
    const dateObj = new Date(start);
    document.getElementById("auto-date").innerText = dateObj.toLocaleDateString('de-DE');
    document.getElementById("auto-time").innerText = dateObj.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  }
  document.getElementById("auto-room").innerText = item.location || "Nicht angegeben";
  document.getElementById("auto-host").innerText = item.organizer || "Unbekannt";
}

async function saveRequirements(event) {
  event.preventDefault();

  // Hilfsfunktion für Checkbox-Arrays (Fix: Klammer ergänzt)
  const getCheckedValues = (name) =>
    Array.from(document.querySelectorAll(`input[name="${name}"]:checked`)).map(cb => cb.value);

  const formData = {
    date: document.getElementById("auto-date").innerText,
    time: document.getElementById("auto-time").innerText,
    room: document.getElementById("auto-room").innerText,
    host: document.getElementById("auto-host").innerText,
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

  // Validierung angepasst an neue Feldnamen
  if (!formData.project || !formData.guestCount) {
    alert("Bitte füllen Sie mindestens das Projekt und die Gästeanzahl aus!");
    return;
  }

  // Validierung Bewirtung: Wenn Catering gewählt wurde, muss mind. eine Sorte angehakt sein
  if (formData.catering.length === 0 && !document.getElementById("cateringRequested")?.checked) {
    // Hier Logik anpassen: Ist Bewirtung Pflicht oder optional?
  }

  console.log("Gesammelte Daten:", formData);
  alert("Daten wurden lokal erfasst!");
}
