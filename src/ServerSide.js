//const ss = SpreadsheetApp.getActiveSpreadsheet()

function getTickerList() {
  const settingsWS = ss.getSheetByName('Settings')

  const tickerCol = ss.getRangeByName('tickers').getColumn()
  const headerRow = ss.getRangeByName('header_settings').getRow()
  const lastRow = settingsWS
    .getRange(settingsWS.getMaxRows(), tickerCol)
    .getNextDataCell(SpreadsheetApp.Direction.UP)
    .getRow()

  const tickerList = settingsWS
    .getRange(headerRow + 1, tickerCol, lastRow - 1)
    .getValues()
    .flat() //Array of array to one dimesion array
  return tickerList
}

function getEstoqueValues() {
  const estoqueWS = ss.getSheetByName('Estoque')
  let initialCelRef = ss
    .getRangeByName('first_header_estoque_cell')
    .getA1Notation()
  const firstRowA1Not = `${initialCelRef[0]}${parseInt(initialCelRef[1]) + 1}`
  const estoqueLastRow = estoqueWS
    .getRange('first_header_estoque_cell')
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow()

  const endCellRef = estoqueWS
    .getRange('first_header_estoque_cell')
    .getNextDataCell(SpreadsheetApp.Direction.NEXT)
    .getA1Notation()

  const estoqueRangeA1Notation =
    firstRowA1Not + ':' + endCellRef[0] + estoqueLastRow

  const estoqueRangeValues = estoqueWS
    .getRange(estoqueRangeA1Notation)
    .getValues()

  return estoqueRangeValues
}

function teste() {
  const arr = getEstoqueValues().map(t => t[0])
  console.log(arr)
}

function getPositionAsset(ticker) {
  if (
    getTickerList().length < 1 ||
    !getTickerList().includes(ticker.toUpperCase())
  ) {
    return 1
  }

  const compra_brl_ws = ss.getSheetByName('Compras_BRL')
  const last_row = compra_brl_ws.getLastRow()
  let compra_brl_data_values = []
  compra_brl_data_values = compra_brl_ws
    .getRange(2, 1, last_row - 1, 9)
    .getValues()

  const ticker_actual_position = getEstoqueValues()
    .filter(p => p[0] == ticker)
    .flat()

  if (ticker_actual_position.length) {
    return ticker_actual_position[2]
  } else {
    const data_values_by_ticker = compra_brl_data_values.filter(
      r => r[3] == ticker
    )
    const array_positions = data_values_by_ticker.map(r => r[8])
    let max_position = Math.max(...array_positions)
    return max_position + 1
  }
}

function getData() {
  let cash = ss.getRangeByName('caixa_brl').getValue()

  const data = {
    cash,
    tickerList: getTickerList(),
    estoqueRangeValues: getEstoqueValues(),
  }
  return data
}

function newBuy(params) {
  const ws = ss.getSheetByName('Compras_BRL')
  //const settingsWS = ss.getSheetByName('Settings')
  const next_id_cel = ss.getRangeByName('control_next_id_compra_brl')
  let next_id_value = next_id_cel.getValue()
  let position = getPositionAsset(params.ticker)

  let date = stringDate(params.date)
  ws.appendRow([
    next_id_value,
    date,
    params.asset,
    params.ticker,
    params.qtd,
    Number(params.price).toLocaleString('pt-BR'),
    Number(params.taxa).toLocaleString('pt-BR'),
    Number(params.total).toLocaleString('pt-BR'),
    position,
  ])

  next_id_cel.setValue(next_id_value + 1)

  ss.toast(`${ws.getName()} - id: ${next_id_value}`, 'Nova compra registrada!')

  return getData()
}

function newSell(params) {
  const ws = ss.getSheetByName('Vendas_BRL')
  const next_id_cel = ss.getRangeByName('control_next_id_venda_brl')
  const next_id_value = next_id_cel.getValue()

  let date = stringDate(params.date)
  ws.appendRow([
    next_id_value,
    date,
    params.asset,
    params.tickerSale,
    params.qtdSale,
    Number(params.price).toLocaleString('pt-BR'),
    Number(params.taxa).toLocaleString('pt-BR'),
    Number(params.totalSale).toLocaleString('pt-BR'),
    Number(params.gainLoss).toLocaleString('pt-BR'),
  ])

  next_id_cel.setValue(next_id_value + 1)
  ss.toast(`${ws.getName()} - id: ${next_id_value}`, 'Nova venda registrada!')

  return getData()
}

function newRecordProvento(params) {
  const proventos_ws = ss.getSheetByName('Proventos_BRL')
  const next_id_field_provento = ss.getRangeByName(
    'control_next_id_proventos_brl'
  )
  const next_id_value = next_id_field_provento.getValue()
  let date = stringDate(params.date)

  console.log('newRecordProvento aqui..');
  proventos_ws.appendRow([
    next_id_value,
    date,
    params.ticker,
    params.type,
    Number(params.value).toLocaleString('pt-BR'),
  ])
  next_id_field_provento.setValue(next_id_value + 1)
  ss.toast(
    `${proventos_ws.getName()} - id: ${next_id_value}`,
    'Novo registro criado!'
  )

  return getData()
}

function newRecordAccount(params) {
  const conta_corrente_ws = ss.getSheetByName('Conta_Corrente')
  const next_id_field_balance = ss.getRangeByName(
    'control_next_id_conta_corrente'
  )
  const next_id_value = next_id_field_balance.getValue()

  let date = stringDate(params.date)
  let operationValue = params.type === 'RETIRADA' ? -params.value : params.value
  conta_corrente_ws.appendRow([
    next_id_value,
    date,
    params.type,
    Number(operationValue).toLocaleString('pt-BR'),
  ])

  next_id_field_balance.setValue(next_id_value + 1)
  ss.toast(
    `${conta_corrente_ws.getName()} - id: ${next_id_value}`,
    'Novo registro criado!'
  )

  return getData()
}

function newRecordDerivative(params) {
  const derivative_ws = ss.getSheetByName('Derivativos PUT_CALL')
  const next_id_field_derivative = ss.getRangeByName(
    'control_next_id_derivative'
  )
  const next_id_value = next_id_field_derivative.getValue()

  let date = stringDate(params.date)
  let deadline = stringDate(params.deadline)
  let operationValue = params.type === 'Compra' ? -params.total*1.0003 : params.total*0.99997

  derivative_ws.appendRow([
    next_id_value,
    date,
    params.type,
    params.derivative,
    params.code,
    deadline,
    params.qtd,
    Number(params.price).toLocaleString('pt-BR'),
    Number(operationValue).toLocaleString('pt-BR'),
  ])

  next_id_field_derivative.setValue(next_id_value + 1)
  ss.toast(
    `${derivative_ws.getName()} - id: ${next_id_value}`,
    'Novo registro criado!'
  )

  return getData()
}

function stringDate(params) {
  let [year, month, day] = params.split('-')
  return `${day}/${month}/${year}`
}
