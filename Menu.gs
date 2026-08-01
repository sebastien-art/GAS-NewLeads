function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Leads")
    .addItem("Procesar Leads", "processNewLeads")
    .addItem("Pasar a Inactivos", "procesarMovimientos")
    .addItem("Aplicar Fórmulas", "aplicarFormulasYRefs_SeleccionSegura")
    .addToUi();
}