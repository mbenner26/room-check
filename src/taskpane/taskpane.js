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
}

/**
 * Sammelt die Daten aus dem Formular und bereitet sie vor
 */
async function saveRequirements(event) {
  event.preventDefault(); // Verhindert das Neuladen der Seite

  // 1. Daten aus dem Formular sammeln
  const formData = {
    presentation: document.getElementById("presentation").value,
    externalGuests: document.getElementById("externalGuests").value,
    guestDevice: document.getElementById("guestDevice").value,
    meetingType: document.getElementById("meetingType").value,
    timestamp: new Date().toISOString()
  };

  // Validierung: Prüfen ob Pflichtfelder ausgefüllt sind
  if (!formData.presentation || !formData.externalGuests || !formData.meetingType) {
    alert("Bitte alle Pflichtfelder ausfüllen!");
    return;
  }

  // Falls externe Gäste "Ja" sind, muss das Gerät gewählt sein
  if (formData.externalGuests === "yes" && !formData.guestDevice) {
    alert("Bitte wählen Sie das Gerät der externen Gäste aus.");
    return;
  }

  console.log("Gesammelte Daten:", formData);

  // 2. Hier findet später der API-Call statt (Phase 2)
  // Wir simulieren hier erst einmal eine erfolgreiche Speicherung
  alert("Daten wurden lokal erfasst! (Vorbereitung für Web-API)");

  // TODO: API Call an Azure Function / Web-Service
  // await postDataToBackend(formData);
}
