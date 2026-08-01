function verReglasCompleto() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const reglas = sheet.getConditionalFormatRules();
  
  reglas.forEach((r, i) => {
    const rangos = r.getRanges().map(rng => rng.getA1Notation()).join(", ");
    const condicion = r.getBooleanCondition();
    const gradiente = r.getGradientCondition();
    
    if (condicion) {
      const criterio = condicion.getCriteriaType();
      const valores = condicion.getCriteriaValues();
      const fondo = condicion.getBackground();
      Logger.log(`Regla ${i+1}: [${rangos}] | ${criterio} | Valores: ${JSON.stringify(valores)} | Fondo: ${fondo}`);
    } else if (gradiente) {
      Logger.log(`Regla ${i+1}: [${rangos}] | Es gradiente`);
    }
  });
}

function repararFormatos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const reglas = [];

  // Regla 1: [A2:AE292] | =$B2="En espera" | #E6B8AF (rosa)
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$B2="En espera"')
    .setBackground('#E6B8AF')
    .setRanges([sheet.getRange('A2:AE292')])
    .build());

  // Regla 2: [P2:P292] | duplicados en P | #F4CCCC (rosa claro)
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($P2<>"",COUNTIF($P:$P,$P2)>1)')
    .setBackground('#F4CCCC')
    .setRanges([sheet.getRange('P2:P292')])
    .build());

  // Regla 3: [V2:V292] | V = "D" | #92D050 (verde)
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('D')
    .setBackground('#92D050')
    .setRanges([sheet.getRange('V2:V292')])
    .build());

  // Regla 4: [V2:V292] | V = "C" | #FF0000 (rojo)
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('C')
    .setBackground('#FF0000')
    .setRanges([sheet.getRange('V2:V292')])
    .build());

  // Regla 5: [E2:E292] | = 0 | #B7E1CD (verde claro)
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBackground('#B7E1CD')
    .setRanges([sheet.getRange('E2:E292')])
    .build());

  // Regla 6: [K2:K292] | < 0 | #F4CCCC (rosa claro)
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#F4CCCC')
    .setRanges([sheet.getRange('K2:K292')])
    .build());

  // Regla 7: [A2:AT292] | estados finales | #999999 (gris)
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($B2="Descalificado",$B2="Perdió interés",$B2="Sin Respuesta",$B2="Ganado")')
    .setBackground('#999999')
    .setRanges([sheet.getRange('A2:AT292')])
    .build());

  // Regla 8: [A2:AJ292] | =$B2="En espera" | #E6B8AF (rosa)
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$B2="En espera"')
    .setBackground('#E6B8AF')
    .setRanges([sheet.getRange('A2:AJ292')])
    .build());

  // Regla 9: [B2:B292] | = 0 | #92D050 (verde)
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBackground('#92D050')
    .setRanges([sheet.getRange('B2:B292')])
    .build());

  // Regla 10: [D2:D292] | duplicados en D | #FF0000 (rojo)
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=COUNTIF($D$2:$D$989,D2)>1')
    .setBackground('#FF0000')
    .setRanges([sheet.getRange('D2:D292')])
    .build());

  // Regla 11: [L1:L292] | GRADIENTE — no se puede restaurar por script
  // Si se rompe, rehacerla manualmente en Formato > Formato condicional

  // Regla 12: [K1:K292] | > 0 | #B7E1CD (verde claro)
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#B7E1CD')
    .setRanges([sheet.getRange('K1:K292')])
    .build());

  // Regla 13: [A2:BB292] | =$B2="Pendiente de propuesta" | #FF9900 (naranja)
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$B2="Pendiente de propuesta"')
    .setBackground('#FF9900')
    .setRanges([sheet.getRange('A2:BB292')])
    .build());

  // Regla 14: [A2:AB292] | =$A2="SD" | #FFFF00 (amarillo)
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2="SD"')
    .setBackground('#FFFF00')
    .setRanges([sheet.getRange('A2:AB292')])
    .build());

  // Regla 15: [A2:AB292] | =$E2="SC" | #B7E1CD (verde claro)
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E2="SC"')
    .setBackground('#B7E1CD')
    .setRanges([sheet.getRange('A2:AB292')])
    .build());

  sheet.setConditionalFormatRules(reglas);
  SpreadsheetApp.getUi().alert('✅ Formatos restaurados correctamente (excepto gradiente de columna L)');
}