function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💼 IEM | Leads")
    .addItem("Procesar Leads", "processNewLeads")
    .addItem("Pasar a Inactivos", "procesarMovimientos")
    .addItem("Aplicar Fórmulas", "aplicarFormulasYRefs_SeleccionSegura")
    .addToUi();
}