const ss = SpreadsheetApp.getActiveSpreadsheet()
const form_ws = ss.getSheetByName('Formularios')

//Conta Corrente

function saveRecordBalance() { //conta corrente
  const id_field_conta_corrente = form_ws.getRange('C2')
  const id_value_conta_corrente = id_field_conta_corrente.getValue()
  //console.log(id_value_conta_corrente)
  ss.getRangeByName('data_conta_corrente').setBackground('white')
  
  const form_cells_conta_corrente = ['C5', 'C6', 'C7']
  if(!validateFields(form_cells_conta_corrente)){
    ss.toast(`Preencha os campos obrigatórios.`, `Erro!!! ${conta_corrente_ws.getName()}`)
    return
  } 

  if(id_value_conta_corrente == ''){
    createNewRecordBalance()
    return
  }
  
  const conta_corrente_ws = ss.getSheetByName('Conta_Corrente')
  const cell_balance_found = conta_corrente_ws.getRange('A:A')
                              .createTextFinder(id_value_conta_corrente)
                              .matchCase(true)
                              .matchEntireCell(true)
                              .findNext()

  if(!cell_balance_found) return


  const row_balance = cell_balance_found.getRow()
//  console.log(row_balance)
  const form_values_conta_corrente = form_cells_conta_corrente.map(f => form_ws.getRange(f).getValue())
  form_values_conta_corrente[1] == 'RETIRADA' ? form_values_conta_corrente[2] = -form_values_conta_corrente[2] : null
  form_values_conta_corrente.unshift(id_value_conta_corrente)
  conta_corrente_ws.getRange(row_balance,1,1,form_values_conta_corrente.length).setValues([form_values_conta_corrente])
  ss.toast(`${conta_corrente_ws.getName()} - id: ${id_value_conta_corrente}`, 'Registro alterado!')
  clearFormContaCorrente()
}

function createNewRecordBalance(){
  const conta_corrente_ws = ss.getSheetByName('Conta_Corrente')
  const form_cells_conta_corrente = ['C5', 'C6', 'C7']
  const id_field_conta_corrente = form_ws.getRange('C2')
  const form_values_conta_corrente = form_cells_conta_corrente.map(f => form_ws.getRange(f).getValue())
  const next_id_field_balance = ss.getRangeByName('control_next_id_conta_corrente')
  const next_id_value = next_id_field_balance.getValue()
  //registra valor da retirada como negativo
  form_values_conta_corrente[1] == 'RETIRADA' ? form_values_conta_corrente[2] = -form_values_conta_corrente[2] : null

  form_values_conta_corrente.unshift(next_id_value)

  //console.log(form_values_conta_corrente)
  conta_corrente_ws.appendRow(form_values_conta_corrente)
  id_field_conta_corrente.setValue(next_id_value)
  next_id_field_balance.setValue(next_id_value +1)
  ss.toast(`${conta_corrente_ws.getName()} - id: ${next_id_value}`, 'Novo registro criado!')
  clearFormContaCorrente()

}

function clearFormContaCorrente(){
  form_ws.getRange('C2').clearContent() //id
  form_ws.getRange('C4').clearContent() //search
  const form_range_conta_corrente = ['C5', 'C6', 'C7']
  form_range_conta_corrente.forEach(f => form_ws.getRange(f).clearContent())
}

function searchBalance(){
  const conta_corrente_ws = ss.getSheetByName('Conta_Corrente')
  const id_field_conta_corrente = form_ws.getRange('C2')
  const search_value_balance = form_ws.getRange('C4').getValue()
  const form_range_conta_corrente = ['C5', 'C6', 'C7']
  const data_balance = conta_corrente_ws.getRange('A2:E').getValues()
  const records_balance_found = data_balance.filter(r => r[4] == search_value_balance)
  if(records_balance_found.length == 0) return
  records_balance_found[0][3] = Math.abs(records_balance_found[0][3]) //reverte valores negativos das retiradas
  id_field_conta_corrente.setValue(records_balance_found[0][0])
  form_range_conta_corrente.forEach((f,i) => form_ws.getRange(f).setValue(records_balance_found[0][i+1]))
}

function deleteRecordBalance(){
  const id_balance_value = form_ws.getRange('C2').getValue()

  if(!id_balance_value) return
  const cell_balance_found = conta_corrente_ws.getRange('A:A')
                              .createTextFinder(id_balance_value)
                              .matchCase(true)
                              .matchEntireCell(true)
                              .findNext()

  if(!cell_balance_found) return
  const form_range_conta_corrente = ['C5', 'C6', 'C7']
  const row_balance = cell_balance_found.getRow()
  conta_corrente_ws.deleteRow(row_balance)
  clearFormContaCorrente(form_range_conta_corrente)
  ss.toast(`${conta_corrente_ws.getName()} - id: ${id_balance_value}`, 'Registro deletado!')

}

//Dollar account
const conta_dolar_ws = ss.getSheetByName('Conta_Dolar')
const form_range_conta_dolar = ['f5','f6','f7','f8','f9'] // f9 = formula total
const search_field_dolar = form_ws.getRange('f4')

function saveRecordDolar() { //conta dolar
  const id_cell_dolar = ss.getRangeByName('id_dolar')
  const id_value_dolar = id_cell_dolar.getValue()

  ss.getRangeByName('data_conta_dolar').setBackground('white')
  
  if(!validateFields(form_range_conta_dolar)){
    ss.toast(`Preencha os campos obrigatórios.`, `Erro!!! ${conta_dolar_ws.getName()}`)
    return
  } 
  if(id_value_dolar == ''){
    createNewRecordDolar()
    return
  }

  const cell_dolar_found = conta_dolar_ws.getRange('A:A')
                              .createTextFinder(id_value_dolar)
                              .matchCase(true)
                              .matchEntireCell(true)
                              .findNext()

  if(!cell_dolar_found) return

  const row_dolar = cell_dolar_found.getRow()
//  console.log(row_dolar)
  const field_values_dolar = form_range_conta_dolar.map(f => form_ws.getRange(f).getValue())
  field_values_dolar.unshift(id_value_dolar)
  conta_dolar_ws.getRange(row_dolar,1,1,field_values_dolar.length).setValues([field_values_dolar])
  search_field_dolar.clearContent()
  ss.toast(`${conta_dolar_ws.getName()} - id: ${id_value_dolar}`, 'Registro alterado!')
}

function createNewRecordDolar(){
 
  const id_cell_dolar = ss.getRangeByName('id_dolar') 
  const field_values_dolar = form_range_conta_dolar.map(f => form_ws.getRange(f).getValue())
  const next_id_field_dolar = ss.getRangeByName('control_next_id_conta_dolar')
  const next_id_dolar = next_id_field_dolar.getValue()

  field_values_dolar.unshift(next_id_dolar)

  //console.log(field_values_dolar)
  conta_dolar_ws.appendRow(field_values_dolar)
  id_cell_dolar.setValue(next_id_dolar)
  next_id_field_dolar.setValue(next_id_dolar +1)
  ss.toast(`${conta_dolar_ws.getName()} - id: ${next_id_dolar}`, 'Novo registro criado!')

}

//BOLETA DE COMPRA BRL

function saveCompraBRL(){
  const compra_brl_ws = ss.getSheetByName('Compras_BRL')
  const form_range_compra_brl = ['I5','I6','I7','I8','I9','I10','I11'] //I7=TICKER, 
  const id_compra_brl_value = ss.getRangeByName('id_compra_brl').getValue()

  ss.getRangeByName('data_compra_brl').setBackground('white')
  if(!validateFields(form_range_compra_brl)){
    ss.toast(`Preencha os campos obrigatórios.`, `Erro!!! ${compra_brl_ws.getName()}`)
    return
  } 
  if(id_compra_brl_value == ''){
    createNewCompraBRL()
    return
  }


  const cell_compra_brl_found = compra_brl_ws.getRange('A:A')
                              .createTextFinder(id_compra_brl_value)
                              .matchCase(true)
                              .matchEntireCell(true)
                              .findNext()

  if(!cell_compra_brl_found) return

  const row_compra_brl = cell_compra_brl_found.getRow()
  
  const field_values_boleta_compra_brl = form_range_compra_brl.map(f => form_ws.getRange(f).getValue())
  field_values_boleta_compra_brl.unshift(id_compra_brl_value)
  compra_brl_ws.getRange(row_compra_brl,1,1,field_values_boleta_compra_brl.length)
    .setValues([field_values_boleta_compra_brl])
  clearBoletaCompraBRL()
  ss.toast(`${compra_brl_ws.getName()} - id: ${id_compra_brl_value}`, 'Registro alterado!')
}

function createNewCompraBRL(){
  const compra_brl_ws = ss.getSheetByName('Compras_BRL')
  const form_range_compra_brl = ['I5','I6','I7','I8','I9','I10','I11'] //I7=TICKER,

  const field_values_boleta_compra_brl = form_range_compra_brl.map((f,i) => {
    if(i == 2){ //ticker uppercase
     return form_ws.getRange(f).getValue().toUpperCase()
    }else{
      return form_ws.getRange(f).getValue()
    }

  })
  
  const last_row = compra_brl_ws.getLastRow()
  let new_position = 1
  const ticker_brl = form_ws.getRange('i7').getValue().toUpperCase()
  
  let compra_brl_data_values = []
  if(last_row ==1){
    compra_brl_data_values = compra_brl_ws.getRange(2,1,last_row,9).getValues()
  }else{
    compra_brl_data_values = compra_brl_ws.getRange(2,1,last_row-1,9).getValues()

  }
  const estoque_ws = ss.getSheetByName('Estoque')
  const estoque_brl_values = estoque_ws.getRange('f3:j50').getValues()
 
  //check if ticker has previously position
  const ticker_exist = compra_brl_data_values.some(t => t[3] == ticker_brl)
  //check if ticker has active position
  const ticker_actual_position = estoque_brl_values.filter(p => p[0] == ticker_brl).flat()
  if(last_row == 1 || !ticker_exist){
    field_values_boleta_compra_brl.push(new_position)
  }else if(ticker_actual_position.length){
    field_values_boleta_compra_brl.push(ticker_actual_position[2])
    // new_position = ticker_actual_position[1]
  }else{
    const data_values_by_ticker = compra_brl_data_values.filter(r => r[3] == ticker_brl)
    const array_positions = data_values_by_ticker.map(r =>r[8])
    let max_position = Math.max(...array_positions)
    field_values_boleta_compra_brl.push(max_position +1)
  }
  
 
  const next_id_field_compra_brl = ss.getRangeByName('control_next_id_compra_brl')
  const next_id_compra_brl = next_id_field_compra_brl.getValue()

  field_values_boleta_compra_brl.unshift(next_id_compra_brl)


  compra_brl_ws.appendRow(field_values_boleta_compra_brl)
  ss.getRangeByName('id_compra_brl').setValue(next_id_compra_brl)
  next_id_field_compra_brl.setValue(next_id_compra_brl +1)
  clearBoletaCompraBRL()
  ss.toast(`${compra_brl_ws.getName()} - id: ${next_id_compra_brl}`, 'Nova compra registrada!')
  
}

function searchCompraBRL(){
  const compra_brl_ws = ss.getSheetByName('Compras_BRL')
  const form_range_compra_brl = ['I5','I6','I7','I8','I9']//exclui os campos q sao formulas
  const search_value_compraBRL = form_ws.getRange('i4').getValue()
  const data_compraBRL = compra_brl_ws.getRange('A2:J').getValues()
  const records_comprasBRL_found = data_compraBRL.filter(r => r[9] == search_value_compraBRL)
  if(records_comprasBRL_found.length == 0) return
  ss.getRangeByName('id_compra_brl').setValue(records_comprasBRL_found[0][0])
  form_range_compra_brl
    .forEach((f,i) => form_ws.getRange(f).setValue(records_comprasBRL_found[0][i+1]))
}

function deleteCompraBRL(){
  const compra_brl_ws = ss.getSheetByName('Compras_BRL')
  const id_compra_brl_value = ss.getRangeByName('id_compra_brl').getValue()

  if(!id_compra_brl_value) return
  const cell_compra_brl_found = compra_brl_ws.getRange('A:A')
                              .createTextFinder(id_compra_brl_value)
                              .matchCase(true)
                              .matchEntireCell(true)
                              .findNext()

  if(!cell_compra_brl_found) return

  const row_compra_brl = cell_compra_brl_found.getRow()
  compra_brl_ws.deleteRow(row_compra_brl)
  clearBoletaCompraBRL()
  ss.toast(`${compra_brl_ws.getName()} - id: ${id_compra_brl_value}`, 'Registro deletado!')

}

function clearBoletaCompraBRL(){
  const fields_to_clear = ['i2','i4','I5','I6','I7','I8','I9']
  fields_to_clear.forEach(f => form_ws.getRange(f).clearContent())
  
}

//BOLETA DE VENDA BRL

function saveVendaBRL(){
  const venda_brl_ws = ss.getSheetByName('Vendas_BRL')
  const form_range_venda_brl = ['L5','L6','L9','L10','L11','L12','L13'] //L6 TICKER, 
  const id_venda_brl_value = ss.getRangeByName('id_venda_brl').getValue()

  form_ws.getRange('l5:l6').setBackground('white')
  form_ws.getRange('l9:l10').setBackground('white')
  if(!validateFields(form_range_venda_brl)){
    ss.toast(`Preencha os campos obrigatórios.`, `Erro!!! ${venda_brl_ws.getName()}`)
    return
  } 
  if(id_venda_brl_value == ''){
    createNewVendaBRL()
    return
  }

  const cell_venda_brl_found = venda_brl_ws.getRange('A:A')
                              .createTextFinder(id_venda_brl_value)
                              .matchCase(true)
                              .matchEntireCell(true)
                              .findNext()

  if(!cell_venda_brl_found) return

  const row_venda_brl = cell_venda_brl_found.getRow()
  
  const field_values_boleta_venda_brl = form_range_venda_brl.map(f => form_ws.getRange(f).getValue())
  field_values_boleta_venda_brl.unshift(id_venda_brl_value)
  venda_brl_ws.getRange(row_venda_brl,1,1,field_values_boleta_venda_brl.length)
    .setValues([field_values_boleta_venda_brl])
  clearBoletaVendaBRL()
  ss.toast(`${venda_brl_ws.getName()} - id: ${id_venda_brl_value}`, 'Registro alterado!')
}

function createNewVendaBRL(){
  const venda_brl_ws = ss.getSheetByName('Vendas_BRL')
  const form_range_venda_brl = ['L5','L6','L9','L10','L11','L12','L13'] //L6 TICKER, 
  const field_values_boleta_venda_brl = form_range_venda_brl.map(f => form_ws.getRange(f).getValue())
  const next_id_field_venda_brl = ss.getRangeByName('control_next_id_venda_brl')
  const next_id_venda_brl = next_id_field_venda_brl.getValue()

  field_values_boleta_venda_brl.unshift(next_id_venda_brl)

  venda_brl_ws.appendRow(field_values_boleta_venda_brl)
  ss.getRangeByName('id_venda_brl').setValue(next_id_venda_brl)
  next_id_field_venda_brl.setValue(next_id_venda_brl +1)
  clearBoletaVendaBRL()
  ss.toast(`${venda_brl_ws.getName()} - id: ${next_id_venda_brl}`, 'Nova venda registrada!')
  
}

function searchVendaBRL(){
  const venda_brl_ws = ss.getSheetByName('Vendas_BRL')
  const form_range_venda_brl = ['l5','l6','l9','l10']//exclui os campos q sao formulas
  const search_value_vendaBRL = form_ws.getRange('l4').getValue()
  const data_venda_brl = venda_brl_ws.getRange('A2:i').getValues()
  const records_vendas_brl_found = data_venda_brl.filter(r => r[8] == search_value_vendaBRL)
  if(records_vendas_brl_found.length == 0) return
  ss.getRangeByName('id_venda_brl').setValue(records_vendas_brl_found[0][0])
  form_range_venda_brl
    .forEach((f,i) => form_ws.getRange(f).setValue(records_vendas_brl_found[0][i+1]))
}



function deleteVendaBRL(){
  const venda_brl_ws = ss.getSheetByName('Vendas_BRL')
  const id_venda_brl_value = ss.getRangeByName('id_venda_brl').getValue()


  const cell_venda_brl_found = venda_brl_ws.getRange('A:A')
                              .createTextFinder(id_venda_brl_value)
                              .matchCase(true)
                              .matchEntireCell(true)
                              .findNext()

  if(!cell_venda_brl_found) return

  form_ws.getRange('l5:l6').setBackground('white')
  form_ws.getRange('l9:l10').setBackground('white')

  const row_venda_brl = cell_venda_brl_found.getRow()

  venda_brl_ws.deleteRow(row_venda_brl)
  clearBoletaVendaBRL()
  ss.toast(`${venda_brl_ws.getName()} - id: ${id_venda_brl_value}`, 'Registro deletado!')

}

function clearBoletaVendaBRL(){
  form_ws.getRange('l5:l6').setBackground('white')
  form_ws.getRange('l9:l10').setBackground('white')
  const fields_to_clear = ['l2','l4','l5','l6','l9','l10']
  fields_to_clear.forEach(f => form_ws.getRange(f).clearContent())
  
}

//PROVENTOS

function saveRecordProvento() { //conta proventos
  const id_cell_proventos = form_ws.getRange('C19')
  const id_value_provento = id_cell_proventos.getValue()
  //console.log(id_value_provento)
  const proventos_ws = ss.getSheetByName('Proventos_BRL')
  ss.getRangeByName('form_cells_provento').setBackground('white')
  
  const form_cells_provento = ['C22', 'C23', 'C24', 'C25']
  if(!validateFields(form_cells_provento)){
    ss.toast(`Preencha os campos obrigatórios.`, `Erro!!! ${proventos_ws.getName()}`)
    return
  } 

  if(id_value_provento == ''){
    createNewRecordProvento()
    return
  }
  
  const cell_proventos_found = proventos_ws.getRange('A:A')
                              .createTextFinder(id_value_provento)
                              .matchCase(true)
                              .matchEntireCell(true)
                              .findNext()

  if(!cell_proventos_found) return


  const row_balance = cell_proventos_found.getRow()
//  console.log(row_balance)
  const form_values_provento = form_cells_provento.map(f => form_ws.getRange(f).getValue())
  form_values_provento.unshift(id_value_provento)
  proventos_ws.getRange(row_balance,1,1,form_values_provento.length).setValues([form_values_provento])
  ss.toast(`${proventos_ws.getName()} - id: ${id_value_provento}`, 'Registro alterado!')
  clearFormProvento()
}

function createNewRecordProvento(){
  const proventos_ws = ss.getSheetByName('Proventos_BRL')
  const form_cells_provento = ['C22', 'C23', 'C24', 'C25']
  const id_cell_proventos = form_ws.getRange('C19')
  const form_values_provento = form_cells_provento.map(f => form_ws.getRange(f).getValue())
  const next_id_field_provento = ss.getRangeByName('control_next_id_proventos_brl')
  const next_id_value = next_id_field_provento.getValue()
  

  form_values_provento.unshift(next_id_value)

  //console.log(form_values_provento)
  proventos_ws.appendRow(form_values_provento)
  id_cell_proventos.setValue(next_id_value)
  next_id_field_provento.setValue(next_id_value +1)
  ss.toast(`${proventos_ws.getName()} - id: ${next_id_value}`, 'Novo registro criado!')
  clearFormProvento()

}

function clearFormProvento(){
  ss.getRangeByName('form_cells_provento').setBackground('white')
  form_ws.getRange('C19').clearContent() //id
  form_ws.getRange('C21').clearContent() //search
  const form_cells_provento = ['C22', 'C23', 'C24', 'C25']
  form_cells_provento.forEach(f => form_ws.getRange(f).clearContent())
}

function searchProventos(){
  const proventos_ws = ss.getSheetByName('Proventos_BRL')
  const id_cell_proventos = form_ws.getRange('C19')
  const search_value_provento = form_ws.getRange('C21').getValue()
  const form_cells_provento = ['C22', 'C23', 'C24', 'C25']
  const data_proventos = proventos_ws.getRange('A2:F').getValues()
  const records_proventos_found = data_proventos.filter(r => r[5] == search_value_provento)
  if(records_proventos_found.length == 0) return
  id_cell_proventos.setValue(records_proventos_found[0][0])
  form_cells_provento.forEach((f,i) => form_ws.getRange(f).setValue(records_proventos_found[0][i+1]))
}

function deleteRecordProvento(){
  const id_provento_value = form_ws.getRange('C19').getValue()

  if(!id_provento_value) return
  const cell_proventos_found = proventos_ws.getRange('A:A')
                              .createTextFinder(id_provento_value)
                              .matchCase(true)
                              .matchEntireCell(true)
                              .findNext()

  if(!cell_proventos_found) return
  const form_cells_provento = ['C22', 'C23', 'C24', 'C25']
  const row_provento = cell_proventos_found.getRow()
  proventos_ws.deleteRow(row_provento)
  clearFormProvento(form_cells_provento)
  ss.toast(`${proventos_ws.getName()} - id: ${id_provento_value}`, 'Registro deletado!')

}

function validateFields(range){
  let validate = true
  range.forEach(f => {
    if(form_ws.getRange(f).isBlank()){
      form_ws.getRange(f).activate()
      form_ws.getRange(f).setBackground('#f0bec1')
      validate = false
    }
  })
  return validate
}

function testeNameRange(){
  console.log(ss.getRangeByName('id_conta_corrente').getValue())
}

