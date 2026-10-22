/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Initialisierung
    document.getElementById("run").onclick = saveRequirements;
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

/**
 * Sammelt die Daten aus dem Formular und bereitet sie vor
 */
async function saveRequirements(event) {
  event.preventDefault();
  
  // Sammle alle aktivierten Bewirtungs-Optionen in einem Array
  const selectedCatering = Array.from(document.querySelectorAll('input[name="cateringType"]:checked'))
                                .map(cb => cb.value);

  const formData = {
    presentation: document.getElementById("presentation").value,
    externalGuests: document.getElementById("externalGuests").value,
    guestDevice: document.getElementById("guestDevice").value,
    meetingType: document.getElementById("meetingType").value,
    cateringRequested: document.getElementById("cateringRequested").checked,
    cateringTypes: selectedCatering, // Array mit z.B. ["fruchtig", "herzhaft"]
    timestamp: new Date().toISOString()
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
}
